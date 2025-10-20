### Caminho backup Storage:

\\192.168.0.36\bkp\VSC\2025\SETOR\TAG\MES ATUAL\ARQUIVO BACKUP

### Caminho backup Dropbox:
/BKP - STORAGE/BKP/VSC/2025/SETOR/TAG/MES/BACKUP

### Setores: 
CFQ, CPHD, CQM, CQ-MP, CTF, LIQ, LVA, PDI, SPEP, SPP, SPPV, STA, TI

### TAGS:
CRG 08-001, CTP 08-002, CTP 08-003, DCM 35-001, DCM 35-002, EAA 08-002-003, ETF 08-002, HPL 08-003, 
HPL 08-004, HPL 08-005, HPL 08-006, HPL 08-007, HPL 08-010, HPL 08-011, HPL 08-012, HPL 08-014, HPL 08-015
IFV 08-003, AMT 08-001, CRG 08-002, ETF 08-003, INX 08-001, RMN 08-002, RMN 08-003, TOC 08-003, TPK 08-002
AUT 09-005, CTP 29-001, CTP 29-002, Data Analyst - LMS XChange, ETF 09-001.002, IHM 09-002, VTK 09-001
IGT 20-004, IGT 20-005, IGT 20-006, IGT 20-007, ENV 07-004 TAM 07-003, IHM 07-001, AUT 10-005, CEA 18-001
CFE 18-002, CMT 18-001, ETF 18-001, ETF 18-002, FRD 18-001, HPL 18-002, HPL 18-003, HPL 18-004, HPL 18-005
HPL 18-006, HPL 18-007, HPL 30-001, SEA 18-003, SEL 18-002, AUT 12-004, AUT 12-005, AUT 12-005, AUT 12-006
AUT 12-006, AUT 33-001, AUT 33-002, ENV 12-001, ENV 12-002, ENV 12-003, ENV 12-004, ENV 12-005, ENV 12-006
ENV 12-007, ENV 33-001, ENV 33-002, GVP 29-008, REA 12-008 REA 12-009, REA 33-001 REA 33-002, REA 33-003 REA 33-004
RTP 12-001, RTP 12-002, SINGLE, STV 33-001, STV 33-002, TAM 12-001, TAM 12-003, TAM 12-004, TAM 12-006, TAM 33-001
STP 17-001, AUT 05-007, AUT 05-008, ENV 05-001, ENV 05-002, ENV 05-003, ENV 05-005, ENV 05-006, IPE 05-004, IPE 05-005
IPE 05-006, IPE 05-007, IPE 05-008, REA 05-007, RTP 05-001, IHM 05-003, IHM 29-001, IHM 29-002, OSR 29-005, OSR 29-009
TOC 29-004, CONE, IDS, KAV, REDES

## Extensões backups: 
.rar, .zip, .lxcs

## INFO:
1. Caminho dos backups:  
\\192.168.0.36\bkp\VSC\2025\("SETOR")\("TAG")\("MES ATUAL")\("PASTA DO DIA")\ARQUIVO BACKUP COMPACTADO (".rar ou .zip")
2. O script coleta informações do backup de acordo com a data de modificação mais recente. EX: Se um backup for feito dia 1 e a data de modificação estiver dia 3 ele ira registrar como dia 3
("Neste caso é bom anexar de forma manual comentario e chamados especificando pq o backup não foi realizado no dia correto")
3. O script consegue percorrer pelas subpastas ("PASTA DO DIA") detectando o backup mais recente dentro do intervalo estabelecido
4. Para o funcionamento ideal é necessario manter as nomenclaturas de SETOR, TAG e responsável de forma padronizada com a planilha
5. A planilha deve seguir a formatação padrão na qual foi configurada
6. Para a verificação é necessário manter a célula vazia, sem caracteres e sem preenchimento de cores
7. O script verifica cada célula vazia e preenche de acordo com os dados, se o backup não for encontrado a célula
será preenchida com a coloração vermelha e com o texto "NOT FOUND"
8. O script ignora verificação ja realizada com as informações do backup e prenchida como verde
9. O script ignora formatação fora do padrão, ou seja, é possivel comentar, mudar texto e preenchimento da célula
10. Células marcada como "NOT FOUND" a cada execução o script deverá realizar a verificação novamente
11. Os backups devem esta compactados e na lista de extensões;
12. Remover itens ocultos (abas, linhas, colunas...) caso contrário ira impactar na verificação
13. Quando o responsavel pelo backup entrar de férias, o verificador deverá modificar a planilha com o nome do Responsavel durante o periodo de férias.
---
## CHECKLIST:
- [x] Enviar e-mail com backups pendentes
- [x] Verificar TAG, SETOR e relacionar com os responsáveis, capturar backups marcados como "NOT FOUND"
- [x] Verificar quantidade de backups em falta
- [x] Adiconar no email texto negrito e cor vermelha para o responsavel
- [x] Adicionar no email texto negrito para a TAG
- [x] Enviar somente 1 email com todas as informações (evitar repetições e adicionar intervalo para envio de emails) - Falta evitar repetições
- [x] Informações do email: "BACKUP AUSENTE" Responsavel, TAG, Setor, Mês, Semana atual, intervalo de dias
- [x] Adicionar na interface opção para marcar  e acionar o envio de emails ou não (Caixa de seleção)
- [ ] Adicionar na interface opção para definir horario de envio de email
- [x] Adicionar icone
- [x] Adicionar .env para variaveis sensiveis
- [ ] Executar como serviço (Rodar a aplicação em segundo plano de maneira automatica)
- [x] Realizar a verificação a cada execução (Manual)
- [ ] Verificação automatica com envio de email 7h da manhã
- [ ] No app ícone de notificações contendo informações de backups não encontrados e seus responsaveis
- [ ] Remover caminho padrão da planilha e deixar somente a seleção via GUI (Deixar padrão a que for selecionado)
- [x] Adicionar parar verificação para cancelamento do envio de email 
- [ ] Adicionar uma limpeza semanal de cache e do Log para não gerar sobrecarga no sistema
- [x] Mudança na interface para a adição de uma barra de carregamento/porcentagem do andamento da verificação de backups
- [x] Reorganização do loadout da interface, troca de menus, redistribuir botões, alteração do menu padrão para um option menu

## ORGANIZAÇÃO DE PASTAS:
Dentro da pasta do mês, criar pasta referente ao dia que foi realizado o backup e colocar dentro da pasta todos os backups feitos dentro do intervalo da semana. Se a semana iniciar dia 1 a 7, colocar todos os backups dentro desse intervalo na mesma pasta.

### Anotações Dropbox:
Token de acess expira após 4h automaticamente
Automação com Selenium falha após encontrar verificação captcha
