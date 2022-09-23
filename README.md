# notificacoes-calibracao-maquinas-solda
## Script de notificação de calibração para máquinas de solda fora da validade


### Problemática:
Notificar via e-mail os responsáveis pela calibração de máquinas de solda quando algum equipamento estiver perto de encerrar a validade, ou que já possua calibração vencida.


### Script Python: (notificacao_calibracao_maquinas_solda.py)
- Extrai as informações de uma tabela preenchida manualmente dentro da rede da empresa;
- Separa as máquinas atrasadas (já vencidas) e para vencer (validade inferior a 30 dias);
- Cria dataframes das colunas TAG (Identificação), Localização e Próxima Calibração;

Caso hajam máquinas com calibração vencida ou para vencer, dispara um e-mail de notificação com ambos os dataframes. O código é executado diariamente.