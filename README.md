RPA Criado para um sistema interno da empresa MOVE3 - Sequoia

Consiste em automatizar a criação de usuários na área de clientes.


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
• Cliente ADM
• Rastreio/TMK
• Rastreio/Consulta

📊 RELATÓRIOS:
Após cada execução, um relatório detalhado é gerado
automaticamente com estatísticas e logs de erro.

🔧 SUPORTE:
Em caso de problemas, verifique:
• Conexão com internet
• Credenciais no arquivo .env
• Formato do arquivo Excel

By. Gustavo Antonelli
