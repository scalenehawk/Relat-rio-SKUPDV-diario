# Painel Diário SKUPDV

O ponto de entrada único é `PAINEL DIÁRIO SKUPDV.bat`, localizado nesta pasta raiz.

## Menu

1. Gera o relatório diário oficial do SKUPDV.
2. Gera as imagens, mensagens e a fila de comunicados a partir do relatório.
3. Abre o painel supervisionado de disparos. O envio continua exigindo confirmação explícita.
0. Fecha o painel.

## Estrutura preservada

- `SKUPDV python`: código e dependências do relatório diário.
- `Disparo SKUPDV aos RNs`: geração dos comunicados, fila, controle e integração supervisionada do WhatsApp.
- `CLIENTES +`, `FATURAMENTO`, `DIAS DE ENTREGA`, cestas e planificadores: fontes operacionais; não devem ser movidos ou apagados.
- `SKUPDV - mês`: relatórios já publicados; são históricos.
- `Historico\Pacotes de backup`: cópias compactadas de segurança dos dois módulos.

## Manutenção segura

O arquivo principal apenas chama os processos existentes por caminhos relativos à pasta raiz. Não coloque regras de cálculo ou envio diretamente no `.bat`; mantenha-as nos respectivos módulos para facilitar testes e auditoria.
