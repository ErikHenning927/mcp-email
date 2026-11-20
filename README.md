# Outlook Web MCP - Triagem Inteligente de E-mails

Servidor MCP para automatizar a leitura, triagem e organização de e-mails corporativos do Microsoft Outlook Web (`outlook.office.com` / `outlook.live.com`) **sem precisar de Azure App Registration ou permissão de administrador**.

## Pastas Configuradas

1. **`entrevista`**:
   - E-mails e convites de calendário referentes a processos seletivos, entrevistas de emprego, vagas, etapas técnicas e bate-papos de recrutamento (ex: *Hubii <> Erik | DevOps*, *Desenvolvedor(a) de Automação e IA | Erick*).
2. **`codigos`**:
   - Códigos de verificação em duas etapas (2FA), OTP, tokens e códigos de login recebidos de plataformas (Riot Games, Google, Steam, Microsoft, GitHub, bancos, etc.).
3. **`garbage`**:
   - Prospecções comerciais frias, ofertas de planos e serviços (Ummense, Copel Seguros/Protege Casa, Freelancer), promoções de streaming (HBO Max, Spotify), newsletters comerciais e propagandas.
4. **`Caixa de Entrada` (Mantidos)**:
   - Faturas de serviços essenciais (ex: Fatura Digital COPEL).
   - Alertas reais de segurança de contas (Google, Microsoft).

## Ferramentas Disponíveis (MCP Tools)

* **`outlook_auto_triage`**: Executa a varredura completa da Caixa de Entrada aplicando as regras automaticamente.
* **`outlook_list_emails`**: Lista e-mails com preview do corpo.
* **`outlook_get_email`**: Lê o corpo completo de um e-mail.
* **`outlook_move_to_folder`**: Move e-mails em lote para uma pasta específica.
* **`outlook_list_folders`**: Lista todas as pastas da caixa postal.
* **`outlook_status`**: Exibe o status da sessão do navegador.

## Como Executar Manualmente via Terminal

```bash
# Rodar a triagem completa de todos os não lidos:
node triage-all.js

# Listar os 10 últimos não lidos da Caixa de Entrada:
node cli.js list 10 unread

# Listar os e-mails na pasta códigos:
node cli.js list 10 codigos

# Listar os e-mails na pasta garbage:
node cli.js list 10 garbage
```
