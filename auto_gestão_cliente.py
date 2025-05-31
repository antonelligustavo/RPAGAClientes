# main.py - Versão Final Completa
import asyncio
import pandas as pd
import logging
import os
from datetime import datetime
from playwright.async_api import async_playwright
from dotenv import load_dotenv
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import sys


ENV_PATH = None  # Será definido dinamicamente

def carregar_env(caminho_env=None):
    """Carrega variáveis de ambiente do arquivo especificado"""
    global ENV_PATH
    
    if caminho_env and os.path.exists(caminho_env):
        ENV_PATH = caminho_env
        load_dotenv(dotenv_path=caminho_env)
        logger.info(f"✅ Arquivo .env carregado: {caminho_env}")
        return True
    else:
        # Tentar carregar .env padrão
        load_dotenv()
        logger.warning("⚠️ Usando .env padrão ou variáveis do sistema")
        return False

# Configurar logging
def configurar_logging():
    """Configura o sistema de logging"""
    log_filename = f'automatizador_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
    
    # Criar formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Configurar logger principal
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # Handler para arquivo
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Handler para console
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger

logger = configurar_logging()

# Configurações do sistema
CONFIG = {
    "url": "https://files.jall.com.br",
    "selectors": {
        "login_frame_pattern": "menu.do",
        "username_field": "#l_username",
        "password_field": "#l_password",
        "login_button": "#entrar",
        "access_link": 'a[href="usuarios_incluiAcesso.do"]',
        "frequency_select": "#frq_id",
        "submit_button": "#enviar",
        "subgroup_select": "#subgrupo",
        "login_gestor": "#loginGestor",
        "email_gestor": "#emailGestor",
        "login_gestor2": "#loginGestor2",
        "email_gestor2": "#emailGestor2",
        "nome": "#nome",
        "obs": "#obs",
        "usuario": "#usuario",
        "email": "#email",
        "filtro_cliente": "#filtro_cliente",
        "tipo_pes_select": 'select[name="tipo_pes_id"]',
        "cargo_select": 'select[name="cargo"]',
        "setor_select": 'select[name="setor"]',
        "lupa_button": 'img[src="imagens/icones/lupa.gif"]',
        "empresa_input": 'input[name="empresa_id"]'
    },
    "values": {
        "frequency_id": "90",
        "subgroup_id": "32",  # Será alterado dinamicamente
        "tipo_pes_id": "1",
        "cargo_id": "55",
        "setor_id": "43",
        "obs_text": "Automatizado pelo RPA",
        "empresa_input_position": 0  # Será alterado dinamicamente
    },
    "timeouts": {
        "navigation": 30000,
        "element": 10000,
        "page_load": 3000,
        "retry_delay": 2
    }
}

class AutomatizadorGestao:
    def __init__(self):
        self.stats = {
            "total": 0,
            "sucessos": 0,
            "erros": 0,
            "usuarios_erro": [],
            "inicio_execucao": None,
            "fim_execucao": None
        }

    async def encontrar_frame(self, page, url_pattern, max_tentativas=15, timeout=1.0):
        """Helper robusto para encontrar frames"""
        logger.debug(f"Procurando frame com padrão: {url_pattern}")
        
        for tentativa in range(max_tentativas):
            try:
                frames = page.frames
                frame = None
                
                for f in frames:
                    if url_pattern in f.url:
                        frame = f
                        break
                
                if frame:
                    logger.debug(f"Frame encontrado na tentativa {tentativa + 1}")
                    return frame
                    
            except Exception as e:
                logger.warning(f"Erro ao procurar frame (tentativa {tentativa + 1}): {e}")
            
            await asyncio.sleep(timeout)
        
        raise RuntimeError(f"Frame com padrão '{url_pattern}' não encontrado após {max_tentativas} tentativas")

    async def aguardar_elemento(self, frame_ou_page, seletor, timeout=15000):
        """Aguarda elemento aparecer na página de forma robusta"""
        try:
            await frame_ou_page.wait_for_selector(seletor, timeout=timeout, state='visible')
            return True
        except Exception as e:
            logger.warning(f"Timeout aguardando elemento {seletor}: {e}")
            return False

    async def fazer_login(self, page):
        """Realiza login no sistema"""
        try:
            logger.info("🔐 Iniciando processo de login...")
            
            await page.goto(CONFIG["url"], wait_until='networkidle', timeout=CONFIG["timeouts"]["navigation"])
            
            # Encontrar frame de login
            frame = await self.encontrar_frame(page, CONFIG["selectors"]["login_frame_pattern"])
            
            # Aguardar campos de login
            if not await self.aguardar_elemento(frame, CONFIG["selectors"]["username_field"]):
                raise Exception("Campo de usuário não encontrado")
            
            # Obter credenciais
            username = os.getenv('APP_USERNAME', 'rpa.gestaoac')
            password = os.getenv('APP_PASSWORD')
            
            if not password:
                raise Exception("Senha não encontrada. Configure APP_PASSWORD no arquivo .env")
            
            # Preencher e submeter login
            await frame.fill(CONFIG["selectors"]["username_field"], username)
            await frame.fill(CONFIG["selectors"]["password_field"], password)
            await frame.click(CONFIG["selectors"]["login_button"])
            
            # Aguardar login ser processado
            await page.wait_for_timeout(CONFIG["timeouts"]["page_load"])
            
            logger.info("✅ Login realizado com sucesso")
            return frame
            
        except Exception as e:
            logger.error(f"❌ Erro no login: {e}")
            raise

    async def navegar_para_incluir_acesso(self, page, frame):
        """Navega para a página de incluir acesso"""
        try:
            logger.debug("🧭 Navegando para incluir acesso...")
            
            # Clicar no link de acesso
            await frame.click(CONFIG["selectors"]["access_link"])
            await page.wait_for_timeout(CONFIG["timeouts"]["page_load"])
            
            # Encontrar novo frame
            target_frame = await self.encontrar_frame(page, "usuarios_incluiAcesso.do")
            
            # Aguardar e configurar frequência
            if await self.aguardar_elemento(target_frame, CONFIG["selectors"]["frequency_select"]):
                await target_frame.select_option(CONFIG["selectors"]["frequency_select"], CONFIG["values"]["frequency_id"])
                await target_frame.click(CONFIG["selectors"]["submit_button"])
                await asyncio.sleep(CONFIG["timeouts"]["retry_delay"])
            else:
                raise Exception("Campo de frequência não encontrado")
            
            logger.debug("✅ Navegação para incluir acesso concluída")
            return target_frame
            
        except Exception as e:
            logger.error(f"❌ Erro na navegação: {e}")
            raise

    async def configurar_grupo(self, page):
        """Configura o grupo do usuário"""
        try:
            logger.debug("👥 Configurando grupo...")
            
            # Encontrar frame de grupo
            target_frame = await self.encontrar_frame(page, "usuarios_incluiGrupo.do")
            
            # Aguardar e configurar subgrupo
            if await self.aguardar_elemento(target_frame, CONFIG["selectors"]["subgroup_select"]):
                await target_frame.select_option(CONFIG["selectors"]["subgroup_select"], CONFIG["values"]["subgroup_id"])
            else:
                raise Exception("Campo de subgrupo não encontrado")
            
            logger.debug("✅ Grupo configurado com sucesso")
            return target_frame
            
        except Exception as e:
            logger.error(f"❌ Erro na configuração do grupo: {e}")
            raise

    async def preencher_dados_usuario(self, frame, dados):
        """Preenche os dados do usuário no formulário"""
        try:
            usuario = dados.get('usuario', 'N/A')
            logger.debug(f"📝 Preenchendo dados do usuário: {usuario}")
            
            # Campos opcionais de gestores
            campos_opcionais = {
                'loginGestor': CONFIG["selectors"]["login_gestor"],
                'emailGestor': CONFIG["selectors"]["email_gestor"],
                'loginGestor2': CONFIG["selectors"]["login_gestor2"],
                'emailGestor2': CONFIG["selectors"]["email_gestor2"]
            }
            
            for campo, seletor in campos_opcionais.items():
                if campo in dados and pd.notna(dados[campo]) and str(dados[campo]).strip():
                    await frame.fill(seletor, str(dados[campo]).strip())
            
            # Campos obrigatórios
            campos_obrigatorios = {
                'nome': CONFIG["selectors"]["nome"],
                'usuario': CONFIG["selectors"]["usuario"], 
                'email': CONFIG["selectors"]["email"],
                'filtro_cliente': CONFIG["selectors"]["filtro_cliente"]
            }
            
            for campo, seletor in campos_obrigatorios.items():
                if campo not in dados or pd.isna(dados[campo]):
                    raise Exception(f"Campo obrigatório '{campo}' não encontrado ou vazio")
                
                valor = str(dados[campo]).strip()
                if not valor:
                    raise Exception(f"Campo obrigatório '{campo}' está vazio")
                
                await frame.fill(seletor, valor)
            
            # Observações
            await frame.fill(CONFIG["selectors"]["obs"], CONFIG["values"]["obs_text"])
            
            logger.debug("✅ Dados do usuário preenchidos")
            
        except Exception as e:
            logger.error(f"❌ Erro no preenchimento: {e}")
            raise

    async def configurar_selects(self, frame):
        """Configura os campos select do formulário"""
        try:
            logger.debug("⚙️ Configurando campos select...")
            
            selects_config = [
                (CONFIG["selectors"]["tipo_pes_select"], CONFIG["values"]["tipo_pes_id"]),
                (CONFIG["selectors"]["cargo_select"], CONFIG["values"]["cargo_id"]),
                (CONFIG["selectors"]["setor_select"], CONFIG["values"]["setor_id"])
            ]
            
            for seletor, valor in selects_config:
                if await self.aguardar_elemento(frame, seletor, timeout=5000):
                    await frame.select_option(seletor, valor)
                else:
                    logger.warning(f"Select {seletor} não encontrado")
            
            logger.debug("✅ Campos select configurados")
            
        except Exception as e:
            logger.error(f"❌ Erro na configuração dos selects: {e}")
            raise

    async def finalizar_cadastro(self, frame):
        """Finaliza o processo de cadastro"""
        try:
            logger.debug("🏁 Finalizando cadastro...")
            
            # Clicar na lupa
            if await self.aguardar_elemento(frame, CONFIG["selectors"]["lupa_button"]):
                await frame.click(CONFIG["selectors"]["lupa_button"])
                await asyncio.sleep(1)
            else:
                raise Exception("Botão lupa não encontrado")
            
            # Aguardar e selecionar empresa
            if await self.aguardar_elemento(frame, CONFIG["selectors"]["empresa_input"]):
                inputs = frame.locator(CONFIG["selectors"]["empresa_input"])
                count = await inputs.count()
                
                posicao = CONFIG["values"]["empresa_input_position"]
                if posicao < count:
                    await inputs.nth(posicao).click()
                else:
                    logger.warning(f"Posição {posicao} não existe, usando posição 0")
                    await inputs.nth(0).click()
            
            await asyncio.sleep(1)
            
            # Executar checkAll
            try:
                await frame.evaluate("checkAll()")
            except Exception as e:
                logger.warning(f"Erro ao executar checkAll: {e}")
            
            # Submeter formulário
            if await self.aguardar_elemento(frame, CONFIG["selectors"]["submit_button"]):
                await frame.click(CONFIG["selectors"]["submit_button"])
                await asyncio.sleep(2)  # Aguardar processamento
            else:
                raise Exception("Botão submit não encontrado")
            
            logger.debug("✅ Cadastro finalizado")
            
        except Exception as e:
            logger.error(f"❌ Erro na finalização: {e}")
            raise

    async def processar_usuario(self, page, dados, frame_inicial):
        """Processa um único usuário completo"""
        usuario = dados.get('usuario', 'USUÁRIO_DESCONHECIDO')
        
        try:
            logger.info(f"🚀 Processando usuário: {usuario}")
            
            # 1. Navegar para incluir acesso
            frame_acesso = await self.navegar_para_incluir_acesso(page, frame_inicial)
            
            # 2. Configurar grupo
            frame_grupo = await self.configurar_grupo(page)
            
            # 3. Preencher dados do usuário
            await self.preencher_dados_usuario(frame_grupo, dados)
            
            # 4. Configurar selects
            await self.configurar_selects(frame_grupo)
            
            # 5. Finalizar cadastro
            await self.finalizar_cadastro(frame_grupo)
            
            logger.info(f"✅ Usuário {usuario} processado com sucesso!")
            self.stats["sucessos"] += 1
            return True
            
        except Exception as e:
            logger.error(f"❌ Erro ao processar {usuario}: {e}")
            self.stats["erros"] += 1
            self.stats["usuarios_erro"].append({
                "usuario": usuario,
                "erro": str(e),
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            return False

    async def gerar_relatorio(self):
        """Gera relatório final de execução"""
        tempo_execucao = None
        if self.stats["inicio_execucao"] and self.stats["fim_execucao"]:
            tempo_execucao = (self.stats["fim_execucao"] - self.stats["inicio_execucao"]).total_seconds()
        
        logger.info("=" * 60)
        logger.info("📊 RELATÓRIO FINAL DE EXECUÇÃO")
        logger.info("=" * 60)
        logger.info(f"📈 Total de usuários: {self.stats['total']}")
        logger.info(f"✅ Sucessos: {self.stats['sucessos']}")
        logger.info(f"❌ Erros: {self.stats['erros']}")
        
        if self.stats['total'] > 0:
            taxa_sucesso = (self.stats['sucessos'] / self.stats['total'] * 100)
            logger.info(f"📊 Taxa de sucesso: {taxa_sucesso:.1f}%")
        
        if tempo_execucao:
            logger.info(f"⏱️ Tempo de execução: {tempo_execucao:.1f} segundos")
        
        if self.stats["usuarios_erro"]:
            logger.info(f"\n❌ Usuários com erro ({len(self.stats['usuarios_erro'])}):")
            for erro in self.stats["usuarios_erro"]:
                logger.info(f"  • {erro['usuario']}: {erro['erro']}")
        
        # Salvar relatório em JSON
        relatorio_arquivo = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(relatorio_arquivo, 'w', encoding='utf-8') as f:
                json.dump({
                    **self.stats,
                    "tempo_execucao_segundos": tempo_execucao
                }, f, ensure_ascii=False, indent=2, default=str)
            
            logger.info(f"💾 Relatório salvo em: {relatorio_arquivo}")
        except Exception as e:
            logger.error(f"❌ Erro ao salvar relatório: {e}")

    async def executar(self, arquivo_excel):
        """Método principal de execução"""
        self.stats["inicio_execucao"] = datetime.now()
        
        try:
            # Validar arquivo
            if not os.path.exists(arquivo_excel):
                raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_excel}")
            
            # Carregar dados
            logger.info(f"📂 Carregando dados: {os.path.basename(arquivo_excel)}")
            try:
                df = pd.read_excel(arquivo_excel)
            except Exception as e:
                raise Exception(f"Erro ao ler arquivo Excel: {e}")
            
            if df.empty:
                raise Exception("Arquivo Excel está vazio")
            
            self.stats["total"] = len(df)
            logger.info(f"📋 {len(df)} usuários carregados para processamento")
            
            # Processar usuários
            async with async_playwright() as p:
                for idx, linha in df.iterrows():
                    logger.info(f"\n{'='*50}")
                    logger.info(f"👤 Usuário {idx + 1}/{len(df)}: {linha.get('usuario', 'N/A')}")
                    logger.info(f"{'='*50}")
                    
                    browser = None
                    try:
                        # Configurar browser
                        browser = await p.chromium.launch(
                            headless=False,
                            args=['--no-sandbox', '--disable-dev-shm-usage']
                        )
                        
                        context = await browser.new_context()
                        page = await context.new_page()
                        
                        # Fazer login
                        frame_inicial = await self.fazer_login(page)
                        
                        # Processar usuário
                        await self.processar_usuario(page, linha, frame_inicial)
                        
                        # Pausa entre usuários
                        await asyncio.sleep(CONFIG["timeouts"]["retry_delay"])
                        
                    except Exception as e:
                        logger.error(f"💥 Erro crítico no usuário {idx + 1}: {e}")
                        self.stats["erros"] += 1
                        self.stats["usuarios_erro"].append({
                            "usuario": linha.get('usuario', f'Linha_{idx + 1}'),
                            "erro": f"Erro crítico: {str(e)}",
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                    
                    finally:
                        if browser:
                            await browser.close()
            
            self.stats["fim_execucao"] = datetime.now()
            await self.gerar_relatorio()
            
        except Exception as e:
            self.stats["fim_execucao"] = datetime.now()
            logger.error(f"💥 Erro crítico na execução: {e}")
            await self.gerar_relatorio()
            raise

class InterfaceAutomatizador:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🤖 Automatizador Gestão de Acessos (Clientes) v2.0")
        self.root.geometry("600x1080")  # Aumentar altura
        self.root.resizable(True, True)
        
        # Configurar ícone se existir
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # Configurar estilo
        self.configurar_estilo()
        
        # Variáveis
        self.tipo_cliente_var = tk.StringVar(value="Cliente ADM")
        self.campo_contrato_var = tk.StringVar(value="1")
        self.arquivo_excel = r"C:\Users\gustavo.ribeiro\Desktop\Python\Automatizador gestão de acessos\usuarios.xlsx"
        self.arquivo_env = None  # NOVA VARIÁVEL
        self.executando = False
        
        self.criar_interface()

    # ========== NOVA FUNÇÃO ==========
    def criar_secao_configuracao_arquivos(self, parent):
        """Cria a seção de configuração de arquivos do sistema"""
        config_arquivos_frame = ttk.LabelFrame(parent, text="🔧 Configuração do Sistema", padding="15")
        config_arquivos_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Arquivo .env
        env_frame = ttk.Frame(config_arquivos_frame)
        env_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            env_frame,
            text="🔐 Arquivo de Configuração (.env):",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.env_label = ttk.Label(
            env_frame,
            text="❌ Nenhum arquivo .env selecionado",
            font=("Arial", 9),
            foreground="red"
        )
        self.env_label.pack(anchor=tk.W, pady=(0, 5))
        
        ttk.Button(
            env_frame,
            text="📁 Selecionar Arquivo .env",
            command=self.selecionar_arquivo_env
        ).pack(anchor=tk.W)
        
        # Separador
        ttk.Separator(config_arquivos_frame, orient='horizontal').pack(fill=tk.X, pady=(10, 15))
        
        # Botão para testar configurações
        ttk.Button(
            config_arquivos_frame,
            text="🧪 Testar Configurações",
            command=self.testar_configuracoes
        ).pack(anchor=tk.W)

    # ========== NOVA FUNÇÃO ==========
    def selecionar_arquivo_env(self):
        """Abre diálogo para selecionar arquivo .env"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo .env",
            filetypes=[
                ("Arquivos ENV", "*.env"),
                ("Todos os arquivos", "*.*")
            ],
            initialdir=os.getcwd()
        )
        
        if arquivo:
            if carregar_env(arquivo):
                self.arquivo_env = arquivo
                self.env_label.config(
                    text=f"✅ {os.path.basename(arquivo)} - Carregado com sucesso",
                    foreground="green"
                )
            else:
                messagebox.showerror("Erro", "Não foi possível carregar o arquivo .env")

    # ========== NOVA FUNÇÃO ==========
    def testar_configuracoes(self):
        """Testa se as configurações estão corretas"""
        try:
            username = os.getenv('APP_USERNAME')
            password = os.getenv('APP_PASSWORD')
            
            if not username:
                messagebox.showerror("Erro", "APP_USERNAME não encontrado no arquivo .env")
                return
            
            if not password:
                messagebox.showerror("Erro", "APP_PASSWORD não encontrado no arquivo .env")
                return
            
            messagebox.showinfo(
                "Configurações OK", 
                f"✅ Configurações carregadas com sucesso!\n\nUsuário: {username}\nSenha: {'*' * len(password)}"
            )
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao testar configurações:\n{e}")
    
    def configurar_estilo(self):
        """Configura o estilo da interface"""
        self.style = ttk.Style()
        
        # Tentar usar tema moderno
        temas_disponiveis = self.style.theme_names()
        
        if 'vista' in temas_disponiveis:
            self.style.theme_use('vista')
        elif 'clam' in temas_disponiveis:
            self.style.theme_use('clam')
        elif 'alt' in temas_disponiveis:
            self.style.theme_use('alt')
    
    def criar_interface(self):
        """Cria a interface gráfica completa"""
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Cabeçalho
        self.criar_cabecalho(main_frame)
        
        # NOVA SEÇÃO - Configuração de arquivos do sistema
        self.criar_secao_configuracao_arquivos(main_frame)
        
        # Seção arquivo
        self.criar_secao_arquivo(main_frame)
        
        # Seção configurações
        self.criar_secao_configuracoes(main_frame)
        
        # Seção controles
        self.criar_secao_controles(main_frame)
        
        # Seção logs
        self.criar_secao_logs(main_frame)
        
        # Configurar logging para interface
        self.configurar_logs_interface()
    
    def criar_cabecalho(self, parent):
        """Cria o cabeçalho da aplicação"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        titulo = ttk.Label(
            header_frame,
            text="🤖 Automatizador Gestão de Acessos",
            font=("Arial", 20, "bold")
        )
        titulo.pack()
        
        subtitulo = ttk.Label(
            header_frame,
            text="Sistema RPA para criação automatizada de usuários",
            font=("Arial", 11),
            foreground="gray"
        )
        subtitulo.pack(pady=(5, 0))
    
    def criar_secao_arquivo(self, parent):
        """Cria a seção de seleção de arquivo"""
        arquivo_frame = ttk.LabelFrame(parent, text="📁 Arquivo Excel", padding="15")
        arquivo_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Frame para arquivo atual
        current_frame = ttk.Frame(arquivo_frame)
        current_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.arquivo_label = ttk.Label(
            current_frame,
            text=f"📄 {os.path.basename(self.arquivo_excel)}",
            font=("Arial", 11, "bold")
        )
        self.arquivo_label.pack(anchor=tk.W)
        
        self.status_label = ttk.Label(current_frame, font=("Arial", 9))
        self.status_label.pack(anchor=tk.W, pady=(2, 0))
        
        self.caminho_label = ttk.Label(
            current_frame,
            text=f"📂 {self.arquivo_excel}",
            font=("Arial", 8),
            foreground="gray"
        )
        self.caminho_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Botão para selecionar arquivo
        ttk.Button(
            arquivo_frame,
            text="📂 Selecionar Outro Arquivo",
            command=self.selecionar_arquivo
        ).pack(anchor=tk.W)
        
        self.atualizar_status_arquivo()
    
    def criar_secao_configuracoes(self, parent):
        """Cria a seção de configurações"""
        config_frame = ttk.LabelFrame(parent, text="⚙️ Configurações", padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Tipo Cliente
        tipo_frame = ttk.Frame(config_frame)
        tipo_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            tipo_frame,
            text="👥 Tipo de Cliente:",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(0, 8))
        
        tipos_cliente = [
            ("Cliente ADM", "Cliente ADM"),
            ("Rastreio/TMK", "Rastreio/TMK"),
            ("Rastreio/Consulta", "Rastreio/Consulta")
        ]
        
        for texto, valor in tipos_cliente:
            ttk.Radiobutton(
                tipo_frame,
                text=texto,
                variable=self.tipo_cliente_var,
                value=valor
            ).pack(anchor=tk.W, pady=2)
        
        # Separador
        ttk.Separator(config_frame, orient='horizontal').pack(fill=tk.X, pady=(10, 15))
        
        # Campo Contrato
        contrato_frame = ttk.Frame(config_frame)
        contrato_frame.pack(fill=tk.X)
        
        ttk.Label(
            contrato_frame,
            text="📋 Campo Contrato:",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(0, 8))
        
        campos_contrato = [
            ("Campo 1 (Padrão)", "1"),
            ("Campo 2", "2"),
            ("Campo 3", "3")
        ]
        
        for texto, valor in campos_contrato:
            ttk.Radiobutton(
                contrato_frame,
                text=texto,
                variable=self.campo_contrato_var,
                value=valor
            ).pack(anchor=tk.W, pady=2)
    
    def criar_secao_controles(self, parent):
        """Cria a seção de controles"""
        controles_frame = ttk.Frame(parent)
        controles_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Botão principal
        self.btn_principal = tk.Button(
            controles_frame,
            text="🚀 INICIAR PROCESSAMENTO",
            command=self.iniciar_processamento,
            font=("Arial", 14, "bold"),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief="raised",
            bd=3,
            cursor="hand2",
            height=2
        )
        self.btn_principal.pack(fill=tk.X, pady=(0, 10), ipady=8)
        
        # Frame para botões secundários
        botoes_sec_frame = ttk.Frame(controles_frame)
        botoes_sec_frame.pack(fill=tk.X)
        
        ttk.Button(
            botoes_sec_frame,
            text="📊 Ver Últimos Relatórios",
            command=self.ver_relatorios
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            botoes_sec_frame,
            text="🔧 Testar Conexão",
            command=self.testar_conexao  
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            botoes_sec_frame,
            text="❓ Ajuda",
            command=self.mostrar_ajuda
        ).pack(side=tk.RIGHT)
    
    def criar_secao_logs(self, parent):
        """Cria a seção de logs"""
        logs_frame = ttk.LabelFrame(parent, text="📋 Logs de Execução", padding="10")
        logs_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame para área de texto e scrollbar
        texto_frame = ttk.Frame(logs_frame)
        texto_frame.pack(fill=tk.BOTH, expand=True)
        
        # Área de texto
        self.log_text = tk.Text(
            texto_frame,
            height=15,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#f8f9fa",
            fg="#333333",
            wrap=tk.WORD,
            padx=10,
            pady=5
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(texto_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Frame para controles dos logs
        controles_logs_frame = ttk.Frame(logs_frame)
        controles_logs_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            controles_logs_frame,
            text="🗑️ Limpar Logs",
            command=self.limpar_logs
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            controles_logs_frame,
            text="💾 Salvar Logs",
            command=self.salvar_logs
        ).pack(side=tk.LEFT, padx=(10, 0))
    
    def configurar_logs_interface(self):
        """Configura o logging para aparecer na interface"""
        # Handler personalizado para a interface
        self.gui_handler = LogHandler(self.log_text)
        self.gui_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        ))
        
        # Adicionar handler ao logger principal
        logger.addHandler(self.gui_handler)
    
    def atualizar_status_arquivo(self):
        """Atualiza o status do arquivo Excel"""
        try:
            if os.path.exists(self.arquivo_excel):
                df = pd.read_excel(self.arquivo_excel)
                self.status_label.config(
                    text=f"✅ {len(df)} usuários encontrados",
                    foreground="green"
                )
            else:
                self.status_label.config(
                    text="❌ Arquivo não encontrado",
                    foreground="red"
                )
        except Exception as e:
            self.status_label.config(
                text=f"❌ Erro ao ler arquivo: {str(e)[:50]}...",
                foreground="red"
            )
    
    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo Excel"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo Excel",
            filetypes=[
                ("Arquivos Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*")
            ],
            initialdir=os.path.dirname(self.arquivo_excel)
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            self.arquivo_label.config(text=f"📄 {os.path.basename(arquivo)}")
            self.caminho_label.config(text=f"📂 {arquivo}")
            self.atualizar_status_arquivo()
    
    def iniciar_processamento(self):
        """Inicia o processamento em thread separada"""
        if self.executando:
            messagebox.showwarning("Aviso", "Processamento já está em andamento!")
            return
        
        # NOVA VALIDAÇÃO - Verificar se .env foi carregado
        if not self.arquivo_env:
            messagebox.showerror("Erro", "Selecione o arquivo .env antes de iniciar!")
            return
        
        if not os.path.exists(self.arquivo_excel):
            messagebox.showerror("Erro", "Arquivo Excel não encontrado!")
            return
        
        # Verificar credenciais
        username = os.getenv('APP_USERNAME')
        password = os.getenv('APP_PASSWORD')
        
        if not username or not password:
            messagebox.showerror("Erro", "Credenciais não encontradas no arquivo .env!\n\nVerifique se APP_USERNAME e APP_PASSWORD estão configurados.")
            return
        
        # Confirmar execução
        resposta = messagebox.askyesno(
            "Confirmar Execução",
            f"Iniciar processamento com:\n\n📄 Excel: {os.path.basename(self.arquivo_excel)}\n🔐 Configurações: {os.path.basename(self.arquivo_env)}\n👤 Usuário: {username}\n\nDeseja continuar?"
        )
        
        if resposta:
            self.executando = True
            self.btn_principal.config(
                text="⏳ EXECUTANDO...",
                state="disabled",
                bg="#ff9800"
            )
            
            # Executar em thread separada
            thread = threading.Thread(target=self.executar_automatizador)
            thread.daemon = True
            thread.start()
    
    def executar_automatizador(self):
        """Executa o automatizador em thread separada"""
        try:
            # Atualizar configurações baseadas na interface
            self.atualizar_configuracoes()
            
            # Criar e executar automatizador
            automatizador = AutomatizadorGestao()
            
            # Executar com asyncio
            asyncio.run(automatizador.executar(self.arquivo_excel))
            
            # Notificar conclusão
            self.root.after(0, self.execucao_concluida, True)
            
        except Exception as e:
            logger.error(f"Erro na execução: {e}")
            self.root.after(0, self.execucao_concluida, False, str(e))
    
    def atualizar_configuracoes(self):
        """Atualiza as configurações baseadas na interface"""
        tipo_cliente = self.tipo_cliente_var.get()
        campo_contrato = self.campo_contrato_var.get()
        
        # Mapear tipos de cliente para subgrupo_id
        mapeamento_subgrupo = {
            "Cliente ADM": "32",
            "Rastreio/TMK": "113", 
            "Rastreio/Consulta": "133"
        }
        
        CONFIG["values"]["subgroup_id"] = mapeamento_subgrupo.get(tipo_cliente, "32")
        CONFIG["values"]["empresa_input_position"] = int(campo_contrato) - 1
        
        logger.info(f"Configurações atualizadas - Tipo: {tipo_cliente}, Campo: {campo_contrato}")
    
    def execucao_concluida(self, sucesso, erro=None):
        """Callback chamado quando a execução termina"""
        self.executando = False
        self.btn_principal.config(
            text="🚀 INICIAR PROCESSAMENTO",
            state="normal",
            bg="#4CAF50"
        )
        
        if sucesso:
            messagebox.showinfo("Sucesso", "Processamento concluído! Verifique os logs para detalhes.")
        else:
            messagebox.showerror("Erro", f"Erro durante a execução:\n\n{erro}")
    
    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def salvar_logs(self):
        """Salva os logs em arquivo"""
        try:
            conteudo = self.log_text.get(1.0, tk.END)
            if not conteudo.strip():
                messagebox.showwarning("Aviso", "Não há logs para salvar!")
                return
            
            arquivo = filedialog.asksaveasfilename(
                title="Salvar logs",
                defaultextension=".txt",
                filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
                initialname=f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            
            if arquivo:
                with open(arquivo, 'w', encoding='utf-8') as f:
                    f.write(conteudo)
                messagebox.showinfo("Sucesso", f"Logs salvos em:\n{arquivo}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar logs:\n{e}")
    
    def ver_relatorios(self):
        """Abre diálogo para visualizar relatórios"""
        try:
            arquivos_relatorio = [f for f in os.listdir('.') if f.startswith('relatorio_') and f.endswith('.json')]
            
            if not arquivos_relatorio:
                messagebox.showinfo("Info", "Nenhum relatório encontrado.")
                return
            
            # Ordenar por data (mais recente primeiro)
            arquivos_relatorio.sort(reverse=True)
            
            # Criar janela de relatórios
            janela_relatorio = tk.Toplevel(self.root)
            janela_relatorio.title("📊 Relatórios de Execução")
            janela_relatorio.geometry("600x400")
            
            # Lista de relatórios
            frame_lista = ttk.Frame(janela_relatorio)
            frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            ttk.Label(frame_lista, text="Relatórios Disponíveis:", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
            
            listbox = tk.Listbox(frame_lista, font=("Consolas", 10))
            listbox.pack(fill=tk.BOTH, expand=True)
            
            for arquivo in arquivos_relatorio[:10]:  # Mostrar apenas os 10 mais recentes
                listbox.insert(tk.END, arquivo)
            
            # Botão para abrir relatório
            ttk.Button(
                frame_lista,
                text="📂 Abrir Relatório Selecionado",
                command=lambda: self.abrir_relatorio(arquivos_relatorio[listbox.curselection()[0]] if listbox.curselection() else None)
            ).pack(pady=(10, 0))
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar relatórios:\n{e}")
    
    def abrir_relatorio(self, arquivo):
        """Abre um relatório específico"""
        if not arquivo:
            messagebox.showwarning("Aviso", "Selecione um relatório!")
            return
        
        try:
            os.startfile(arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir relatório:\n{e}")
    
    def testar_conexao(self):
        """Testa a conexão com o sistema"""
        def testar():
            try:
                logger.info("🔧 Testando conexão...")
                
                # Simular teste básico
                import requests
                response = requests.get(CONFIG["url"], timeout=10)
                
                if response.status_code == 200:
                    self.root.after(0, lambda: messagebox.showinfo("Sucesso", "✅ Conexão OK!\nSistema acessível."))
                else:
                    self.root.after(0, lambda: messagebox.showwarning("Aviso", f"⚠️ Status HTTP: {response.status_code}"))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Erro", f"❌ Erro de conexão:\n{e}"))
        
        thread = threading.Thread(target=testar)
        thread.daemon = True
        thread.start()
    
    def mostrar_ajuda(self):
        """Mostra janela de ajuda"""
        ajuda_texto = """
🤖 AUTOMATIZADOR GESTÃO DE ACESSOS v2.0

📋 COMO USAR:

1. CONFIGURAÇÃO INICIAL:
   • Selecione o arquivo .env com suas credenciais
   • Teste as configurações para verificar se estão corretas

2. PREPARE O ARQUIVO EXCEL com as colunas obrigatórias:
   • Login gestor 1 - Login do primeiro gestor
   • Email gestor 1 - E-mail do primeiro gestor
   • Login gestor 2 - Login do segundo gestor
   • Email gestor 2 - E-mail do segundo gestor
   • Nome - Nome completo do usuário
   • Login - Login do usuário
   • E-mail - E-mail do usuário  
   • Cliente - Nome/código do cliente

3. IMPORTANTE: Todas as 8 colunas são obrigatórias e devem estar
   nesta ordem exata na planilha Excel.

4. Configure o tipo de cliente e campo contrato
5. Clique em "INICIAR PROCESSAMENTO"

🔐 ARQUIVO .ENV:
Deve conter:
APP_USERNAME=seu_usuario
APP_PASSWORD=sua_senha

⚙️ CONFIGURAÇÕES:
• Cliente ADM: Usuários administrativos
• Rastreio/TMK: Usuários de rastreamento/telemarketing
• Rastreio/Consulta: Usuários de consulta

📊 RELATÓRIOS:
Após cada execução, um relatório detalhado é gerado
automaticamente com estatísticas e logs de erro.

🔧 SUPORTE:
Em caso de problemas, verifique:
• Conexão com internet
• Credenciais no arquivo .env
• Formato do arquivo Excel

By. Gustavo Antonelli
        """
        
        janela_ajuda = tk.Toplevel(self.root)
        janela_ajuda.title("❓ Ajuda - Automatizador")
        janela_ajuda.geometry("600x550")  # Aumentar altura
        janela_ajuda.resizable(False, False)
        
        texto_ajuda = tk.Text(
            janela_ajuda,
            wrap=tk.WORD,
            padx=20,
            pady=20,
            font=("Arial", 10),
            state=tk.DISABLED
        )
        texto_ajuda.pack(fill=tk.BOTH, expand=True)
        
        texto_ajuda.config(state=tk.NORMAL)
        texto_ajuda.insert(tk.END, ajuda_texto)
        texto_ajuda.config(state=tk.DISABLED)
    
    def executar(self):
        """Executa a interface"""
        self.root.mainloop()

class LogHandler(logging.Handler):
    """Handler personalizado para exibir logs na interface"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        """Emite o log na interface gráfica"""
        try:
            msg = self.format(record)
            
            def append_log():
                self.text_widget.config(state=tk.NORMAL)
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)  # Scroll automático
                self.text_widget.config(state=tk.DISABLED)
            
            # Executar na thread principal
            self.text_widget.after(0, append_log)
            
        except Exception:
            self.handleError(record)

def main():
    """Função principal"""
    try:
        logger.info("🚀 Iniciando Automatizador Gestão de Acessos v2.0")
        
        # Verificar dependências críticas
        try:
            import playwright
        except ImportError:
            print("❌ Playwright não instalado. Execute: pip install playwright")
            print("   Depois execute: playwright install")
            return
        
        # REMOVER verificação automática do .env - agora será feita via interface
        logger.info("📝 Configure o arquivo .env através da interface")
        
        # Executar interface
        app = InterfaceAutomatizador()
        app.executar()
        
    except KeyboardInterrupt:
        logger.info("⏹️ Execução interrompida pelo usuário")
    except Exception as e:
        logger.error(f"💥 Erro crítico: {e}")
        messagebox.showerror("Erro Crítico", f"Erro inesperado:\n\n{e}")
    finally:
        logger.info("🏁 Automatizador finalizado")

if __name__ == "__main__":
    main()