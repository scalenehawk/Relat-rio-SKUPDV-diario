# Gerador do relatório SKU/PDV

O programa `gerar_relatorio_skupdv.py` executa a análise e gera o relatório oficial com as abas `Sugestões`, `Conciliação` e `Base das Sugestões`.

## Uso pelo menu

Abra `Executar Relatorio SKUPDV Python.cmd`, informe a data e responda se o relatório Anala deve ser considerado.

## Uso pelo PowerShell

Sem Anala:

```powershell
python -X utf8 .\gerar_relatorio_skupdv.py --data 20-10-2026 --sem-anala
```

Com o Anala mais recente disponível até a data informada:

```powershell
python -X utf8 .\gerar_relatorio_skupdv.py --data 20-10-2026 --usar-anala
```

Execução não interativa:

```powershell
python -X utf8 .\gerar_relatorio_skupdv.py --data 20-10-2026 --sem-anala --nao-interativo
```

## Comportamento de segurança

- As fontes são selecionadas pela maior data disponível que não ultrapasse a data da análise.
- O programa mostra todos os arquivos selecionados antes do processamento.
- Se o relatório do dia já existir, uma nova versão `_v2`, `_v3` e assim por diante será criada.
- A substituição só ocorre quando `--sobrescrever` é informado explicitamente.
- Valores e preços ausentes permanecem vazios; não são convertidos em zero.
- O arquivo final é reaberto e validado antes da confirmação de sucesso.
