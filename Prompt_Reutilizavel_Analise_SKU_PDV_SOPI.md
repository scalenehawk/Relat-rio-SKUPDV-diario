
Refaça integralmente a análise do KPI SKU/PDV SOPI utilizando sempre os arquivos mais recentes disponíveis neste projeto.

Não quero uma automação, um modelo vazio ou somente uma explicação. Quero o relatório Excel completo, preenchido, conferido, visualmente validado e salvo no caminho definido neste prompt.

## Ação obrigatória acrescentada pelo usuário

> Quero que refaça a análise completa de SKU/PDV, gerando um novo relatório em Excel. Porém, em vez de apresentar uma quantidade de sugestões igual ao GAP, quero a relação de todos os itens da cesta que ainda não foram comprados.

> Para formar essa relação, considere somente os produtos da cesta SOPI nacional que tenham sido comercializados pela revenda em 2026, conforme os relatórios de faturamento. Produtos da cesta nacional sem comercialização pela revenda em 2026 continuam válidos para a apuração oficial do KPI, mas não podem aparecer como sugestão de compra.

> Na aba `Sugestões`, acrescente a coluna `Preço bairro` entre `Produto da cesta não comprado` e `Dias de entrega disponíveis`. Preencha o preço pelo cruzamento do `SKU não comprado` com a primeira coluna da fonte mais recente da pasta `TABELA DE PREÇOS`, utilizando como valor a coluna `VALOR` localizada na coluna J. Esta inclusão não autoriza alterações nas demais colunas ou abas.

Esta instrução deve ser aplicada em toda nova reanálise deste projeto.

Interpretação obrigatória desta ação:

- o GAP deve continuar sendo calculado e apresentado;
- o GAP não deve limitar a quantidade de produtos exibidos;
- a cesta SOPI nacional vigente continua sendo a referência oficial para contar o KPI e reconciliar Ambev, Anala e faturamento;
- para cada PDV que não esteja na condição `Sem ação — 3 fontes convergentes em 10 ou mais`, devem ser listados todos os SKUs distintos da cesta regional elegível para sugestões que ainda não foram comprados;
- a cesta regional elegível para sugestões deve conter somente SKUs que pertençam à cesta SOPI nacional e tenham pelo menos um faturamento elegível da revenda em 2026;
- devem ser removidos da relação os SKUs encontrados no Anala ou no faturamento elegível do mês vigente;
- cada combinação `PDV + SKU ainda não comprado` deve ocupar uma linha;
- cada linha deve trazer o `Preço bairro` correspondente ao SKU quando houver correspondência única e válida na tabela de preços;
- o histórico serve para ordenar os itens, mas não para reduzir a relação;
- produtos sem histórico do próprio PDV também devem permanecer na relação quando fizerem parte da cesta regional elegível para sugestões;
- nenhum SKU específico deve receber tratamento especial: a ação abrange toda a cesta regional elegível para sugestões.

### Sequência obrigatória da ação

Execute a separação entre cesta nacional e cesta regional nesta ordem, sem misturar as finalidades:

1. construa a cesta SOPI nacional vigente aplicando as regras oficiais de elegibilidade;
2. use a cesta nacional para apurar e reconciliar `SKUs Ambev`, `SKUs Anala` e `SKUs Faturamento do mês`;
3. percorra todos os faturamentos elegíveis da revenda disponíveis em 2026;
4. forme a cesta regional pela interseção entre a cesta nacional e os SKUs faturados pela revenda em 2026;
5. identifique separadamente os PDVs com ação comercial, ação comercial com validação e validação com relação preventiva;
6. para cada PDV com ação, forme a união dos SKUs encontrados no Anala e no faturamento elegível do mês vigente;
7. restrinja essa união aos produtos da cesta regional;
8. subtraia a união das compras atuais da cesta regional;
9. grave todos os itens restantes exclusivamente nas colunas `SKU não comprado` e `Produto da cesta não comprado` da aba `Sugestões`;
10. utilize o histórico anterior somente para ordenar as linhas resultantes;
11. não crie uma aba específica para a cesta regional;
12. não altere a cesta utilizada na medição oficial do KPI.

### Exemplos obrigatórios de interpretação

#### Exemplo 1 — formação da cesta regional

Considere uma cesta SOPI nacional simplificada com os SKUs:

`{1001, 1002, 1003, 1004, 1005, 1006}`

Considere que os faturamentos elegíveis da revenda em 2026 contenham:

`{1001, 1002, 1004, 1005, 9000}`

O resultado correto será:

`Cesta regional = {1001, 1002, 1004, 1005}`

Interpretação:

- `9000` não entra porque não pertence à cesta SOPI nacional;
- `1003` e `1006` continuam pertencendo à cesta nacional e continuam válidos para a medição do KPI;
- `1003` e `1006` não aparecem como produtos sugeridos porque não possuem faturamento elegível da revenda em 2026;
- a cesta nacional não deve ser reduzida para quatro itens na conciliação do KPI.

#### Exemplo 2 — relação não limitada pelo GAP

Utilizando a cesta regional `{1001, 1002, 1004, 1005}`, considere um PDV com:

- GAP igual a `1`;
- produtos no Anala: `{1001}`;
- produtos no faturamento elegível do mês: `{1002}`.

A união das compras atuais será:

`{1001, 1002}`

A relação completa dos itens regionais ainda não comprados será:

`{1004, 1005}`

O resultado correto na aba `Sugestões` será duas linhas, apesar de o GAP ser apenas `1`:

- uma linha para o SKU `1004`;
- uma linha para o SKU `1005`.

Não selecione apenas um dos produtos para igualar a quantidade de linhas ao GAP.

#### Exemplo 3 — produto encontrado apenas no faturamento vigente

Se o SKU `1004` não aparecer no Anala, mas aparecer no faturamento elegível do mês vigente para o PDV, considere que existe evidência atual de compra e retire o SKU `1004` da relação de itens não comprados.

O produto não deve ser sugerido novamente somente porque ainda não apareceu no Anala.

#### Exemplo 4 — produto sem histórico do próprio PDV

Se o SKU `1005` pertence à cesta regional porque foi comercializado pela revenda em 2026, mas nunca foi comprado pelo PDV analisado, ele continua na relação de itens não comprados. Classifique e ordene o produto como popular na base SOPI ou como item sem histórico, conforme os dados disponíveis.

#### Exemplo 5 — regra de bloqueio da relação de itens

Considere os seguintes resultados para o mesmo PDV:

- Ambev: `10` SKUs;
- Anala: `8` SKUs;
- faturamento vigente: `9` SKUs.

O resultado correto é manter o PDV com ação comercial e validação, porque as três fontes não concordam que a meta foi atingida. O fato de somente a Ambev registrar 10 não encerra a ação.

A relação de itens ainda não comprados somente deve ser bloqueada quando todas as condições abaixo forem atendidas simultaneamente:

1. Ambev possui contagem disponível;
2. Anala possui contagem disponível;
3. faturamento do mês possui contagem disponível;
4. as três contagens são exatamente iguais;
5. o valor comum é maior ou igual a `10`.

Exemplo sem ação:

- Ambev: `10`;
- Anala: `10`;
- faturamento: `10`.

Exemplo que continua em validação com relação preventiva:

- Ambev: `10`;
- Anala: `11`;
- faturamento: `11`.

Embora todas as fontes estejam acima da meta, os números não coincidem. O GAP será zero, mas o PDV continuará sinalizado para validação e receberá a relação preventiva de itens não comprados. Não interprete essa relação como prova de que o PDV está abaixo da meta.

## 1. Objetivo da análise

Confronte as três leituras do mês vigente:

1. planificador SKU/PDV da Ambev, que representa a fonte da auditoria;
2. planificador Anala, extraído do ERP interno;
3. faturamento do mês vigente, utilizado como contravalidação das compras registradas.

O relatório deve:

- identificar todos os PDVs auditados com comodato SOPI;
- comparar, por PDV, a quantidade de SKUs distintos elegíveis informada pela Ambev, pelo Anala e pelo faturamento do mês;
- classificar convergências, divergências e ausências entre as fontes;
- identificar os PDVs que precisam de ação comercial, validação ou relação preventiva;
- apresentar, para cada PDV que não possua confirmação convergente das três fontes em 10 ou mais, a relação completa de todos os itens distintos da cesta regional elegível para sugestões que ainda não foram comprados no mês vigente;
- utilizar o histórico de faturamento dos meses anteriores somente para ordenar e priorizar essa relação completa;
- acrescentar setor, visitas do VDE e dias disponíveis para entrega;
- preservar as fontes e os relatórios anteriores.

Importante: a quantidade de itens relacionados não deve ser limitada ao GAP. O GAP continua sendo apresentado como indicador de quantos SKUs faltam para o PDV atingir a meta de 10 na cesta SOPI nacional, mas a aba de itens não comprados deve apresentar toda a cesta regional elegível para sugestões ainda disponível para aquele PDV.

## 2. Caminho oficial do projeto

Use como raiz do projeto:

`C:\Users\matheus.manoel\OneDrive - Vyas\Matheus - NS - Sinop\SPO - Gestão A Vista\SKU-PDV`

Procure as fontes mais recentes dentro destas pastas:

- `PLANIFICADOR ANALA`;
- `PLANIFICADOR SKU-PDV`;
- `CESTA SKUPDV`;
- `CLIENTES +`;
- `FATURAMENTO`;
- `DIAS DE ENTREGA`;
- `TABELA DE PREÇOS`.

Não utilize arquivos de outras pastas como fonte principal quando existir uma versão mais recente nos caminhos acima.

Não altere, mova, renomeie ou sobrescreva nenhum arquivo-fonte.

Não crie, consulte, atualize ou registre nada no Obsidian sobre este projeto.

## 3. Inventário obrigatório das fontes

Antes de calcular qualquer resultado:

1. liste os arquivos candidatos encontrados em cada pasta;
2. identifique o arquivo mais recente por data de modificação e pela data indicada no nome ou conteúdo;
3. registre o caminho completo, a data e o horário de cada fonte escolhida;
4. confirme o mês de referência de cada arquivo;
5. confirme a data máxima disponível no faturamento do mês vigente;
6. confirme a data de corte do planificador Anala;
7. confirme a data de corte do planificador Ambev;
8. confirme se Anala, Ambev e faturamento representam cortes comparáveis;
9. liste todos os arquivos de faturamento de 2026 disponíveis, de janeiro até o mês vigente;
10. confirme quais meses de 2026 estão cobertos e registre qualquer mês ausente, arquivo parcial ou diferença de corte que possa reduzir indevidamente a cesta regional de sugestões.
11. identifique a fonte mais recente da pasta `TABELA DE PREÇOS`, confirme a aba utilizada, valide que a primeira coluna contém o código do SKU e que a coluna J possui o cabeçalho `VALOR`;
12. conte SKUs únicos, linhas inválidas, duplicidades e conflitos de preço antes de preencher o relatório.

Se os cortes forem diferentes, não classifique automaticamente toda diferença como erro. Informe na resposta final que parte das divergências pode ser causada pela diferença de atualização entre as fontes.

Sempre prefira fontes do mesmo dia. Quando isso não for possível, use os arquivos mais recentes e registre claramente a diferença de corte.

## 4. Regra oficial do KPI

Aplicar as seguintes regras:

- O KPI é mensal.
- O cálculo é realizado por PDV, e não por quantidade de equipamentos.
- Um PDV com um ou vários comodatos SOPI possui uma única meta mensal.
- O PDV é aderente quando compra 10 ou mais SKUs distintos elegíveis da cesta SOPI dentro do mês vigente.
- A aderência de negócio é definida pelo patamar de 10 SKUs, mas, neste relatório conciliado, a condição operacional `Sem ação` somente pode ser confirmada quando Ambev, Anala e faturamento possuem contagens disponíveis, exatamente iguais e maiores ou iguais a 10.
- Considere somente os SKUs participantes da cesta SOPI oficial mais recente da pasta `CESTA SKUPDV`.
- Considere como elegíveis os produtos classificados como `CERVEJA` ou `MATCH` quando essas categorias existirem na cesta oficial.
- Não crie produtos `MATCH` por inferência quando a cesta oficial atual não possuir essa categoria.
- Exclua Marketplace quando houver coluna ou marcação que identifique o canal, além de Chopp, Bag, Barril, BIB, recipientes Petainer de Chopp e qualquer produto que não pertença à cesta elegível.
- Não exclua automaticamente cervejas em lata ou garrafa somente porque o nome comercial contém a palavra `CHOPP`, como `BRAHMA CHOPP LATA`. O expurgo deve identificar Chopp de fato pela descrição e/ou embalagem.
- Conte SKUs distintos. Quantidade de caixas, unidades ou notas não multiplica a contagem do SKU.
- Restrinja o universo da conciliação aos PDVs SOPI presentes no planificador Ambev.
- Não use toda a base do Anala como universo auditado, porque o Anala contém toda a revenda.

## 5. Normalização obrigatória dos códigos

Antes de qualquer cruzamento:

- trate códigos de PDV e SKU como identificadores, não como valores decimais;
- remova espaços antes e depois do código;
- remova zeros à esquerda apenas para o cruzamento;
- remova sufixos artificiais como `.0`;
- não confunda código do PDV com CNPJ, código de vendedor ou outro identificador;
- não confunda código do produto com descrição ou código de embalagem;
- preserve o código normalizado como texto no Excel final;
- elimine duplicidades da cesta pelo código normalizado do SKU;
- elimine duplicidades de compras usando a combinação `PDV + SKU` para a contagem distinta.

Se um mesmo código de SKU aparecer mais de uma vez na cesta, utilize uma única ocorrência e preserve a descrição válida mais completa.

## 6. Construção da cesta SOPI nacional

Na fonte mais recente da pasta `CESTA SKUPDV`:

1. identifique o arquivo mais recente cujo conteúdo represente explicitamente a cesta SOPI, e não apenas um cadastro geral de produtos;
2. localize as colunas de código, nome e tipo do produto;
3. quando o arquivo possuir uma coluna de identificação da cesta, mantenha somente `SOPI`;
4. quando a própria aba ou arquivo já estiver identificado como `CESTA SOPI`, aceite a aba como o universo SOPI mesmo que não exista uma coluna adicional chamada `TIPO` ou `SOPI`;
5. mantenha somente `CERVEJA` e `MATCH` quando presentes na fonte oficial;
6. não recupere, complete ou recrie a categoria `MATCH` a partir de arquivos antigos se ela não constar na cesta oficial mais recente;
7. exclua Chopp de fato e embalagens `BAG`, `BARRIL`, `BIB` e recipientes `PETAINER` de Chopp;
8. não exclua `BRAHMA CHOPP LATA`, `BRAHMA CHOPP GARRAFA` ou produto equivalente quando a embalagem demonstrar que se trata de cerveja comum e não Chopp faturado;
9. exclua Marketplace quando a fonte possuir marcação confiável para isso;
10. normalize os códigos;
11. elimine códigos vazios;
12. elimine duplicidades de SKU pelo código normalizado;
13. quando o mesmo SKU aparecer em mais de uma categoria, conte-o uma única vez e registre a inconsistência cadastral;
14. calcule a quantidade final de SKUs distintos elegíveis;
15. registre essa quantidade, a composição por categoria, os expurgos e as duplicidades na resposta final.

### 6.0.1. Ponto de controle da cesta vigente em 28/08/2026

Na data desta revisão, a fonte oficial mais recente é:

`CESTA SKUPDV\Nova Cesta Produtos.xlsx`

Estrutura validada:

- aba: `CESTA SOPI`;
- intervalo utilizado: `A1:C4543`;
- colunas: `Cod`, `Nome Produto` e `Tipo Produto`;
- 4.542 registros de produtos;
- 4.542 registros classificados como `CERVEJA`;
- nenhum registro classificado como `MATCH` nessa versão;
- 591 registros expurgados por Chopp ou embalagem de Chopp;
- 3.951 SKUs distintos elegíveis após os expurgos;
- nenhuma duplicidade entre os SKUs elegíveis.

Esses números são pontos de controle da versão de 28/08/2026, e não constantes permanentes. Em cada nova análise, recalcule os totais usando a cesta oficial mais recente. Se os números mudarem, use o resultado recalculado e registre a diferença na resposta final.

Não volte a utilizar automaticamente a antiga cesta `SOPI (1).xlsx` quando existir uma cesta oficial mais recente. Não utilize o cadastro geral `Nova Cesta Produtos (2).xlsx` da pasta Downloads como cesta oficial do projeto.

A cesta mais recente é a referência nacional para decidir se um produto participa da apuração oficial do KPI. Um produto que não pertence à cesta SOPI nacional não deve contar no KPI nem aparecer na relação de itens ainda não comprados.

### 6.1. Construção da cesta regional elegível para sugestões

Esta regra altera somente o universo das sugestões comerciais. Ela não altera a cesta nacional usada na apuração e na reconciliação do KPI.

Utilize todos os relatórios de faturamento disponíveis do ano de 2026, de janeiro até a data de corte do mês vigente. Não use somente o histórico do próprio PDV e não restrinja esta etapa aos PDVs SOPI auditados: o objetivo é identificar quais produtos da cesta são efetivamente comercializados pela revenda na região.

Para identificar os SKUs comercializados pela revenda em 2026:

1. reúna todos os arquivos mensais de faturamento de 2026 disponíveis na pasta `FATURAMENTO`;
2. considere somente operações `1` ou `2`;
3. considere somente status `A` ou `E`;
4. normalize os códigos dos produtos;
5. mantenha somente produtos pertencentes à cesta SOPI nacional vigente;
6. considere o SKU comercializado quando existir pelo menos um registro de faturamento elegível em 2026;
7. elimine duplicidades de SKU;
8. registre os meses efetivamente cobertos, a data máxima do mês vigente e qualquer lacuna encontrada.

Calcule obrigatoriamente:

`Cesta regional elegível para sugestões = Cesta SOPI nacional vigente ∩ SKUs comercializados pela revenda em 2026`

Regras de interpretação:

- basta um faturamento elegível em qualquer mês de 2026 para o SKU entrar na cesta regional de sugestões;
- o faturamento do mês vigente também participa da identificação dos produtos comercializados pela revenda em 2026;
- um SKU da cesta nacional sem faturamento elegível da revenda em 2026 deve ser excluído das abas `Sugestões` e `Base das Sugestões`;
- essa exclusão não retira o SKU da cesta nacional e não muda a quantidade oficial de SKUs comprados informada ou apurada para o KPI;
- se um PDV tiver comprado no mês um SKU da cesta nacional sem comercialização anterior registrada em 2026, esse SKU continua contando normalmente para o KPI, embora não seja uma opção pendente de sugestão;
- não substitua a cesta nacional pela cesta regional nas contagens de `SKUs Ambev`, `SKUs Anala` e `SKUs Faturamento do mês`;
- registre na resposta final a quantidade da cesta nacional, a quantidade da cesta regional elegível para sugestões e a quantidade de SKUs nacionais excluídos das sugestões por ausência de comercialização em 2026.

Não exiba a cesta regional em uma aba adicional. Ela é uma regra intermediária de filtro e deve aparecer no relatório somente por meio:

- dos totais e da explicação metodológica apresentados na resposta final;
- dos produtos resultantes nas colunas `SKU não comprado` e `Produto da cesta não comprado` da aba `Sugestões`;
- dos cálculos resumidos por PDV na aba `Base das Sugestões`.

## 7. Universo dos PDVs auditados

Utilize o planificador SKU/PDV mais recente da Ambev para formar o universo oficial de PDVs com SOPI.

Para cada PDV, capture no mínimo:

- Setor ou código do RN;
- Código do PDV;
- Nome do PDV;
- quantidade de SKUs distintos apontada pela Ambev.

Cada PDV deve aparecer uma única vez na aba `Conciliação`.

Se o planificador possuir mais de uma linha para o mesmo PDV, consolide corretamente por PDV sem multiplicar a meta por quantidade de equipamentos.

## 8. Leitura do planificador Anala

Utilize o planificador Anala mais recente do mês vigente.

Considere as abas que identificam quem comprou e quem não comprou produtos SOPI. Para cada PDV auditado:

- identifique o código do PDV;
- identifique o nome do PDV;
- capture todos os códigos de produtos SOPI comprados no mês;
- normalize os códigos dos produtos;
- mantenha somente os produtos que pertencem à cesta elegível mais recente;
- elimine duplicidades de `PDV + SKU`;
- calcule a quantidade de SKUs distintos elegíveis no Anala.

Um PDV do planificador Ambev que não aparecer no recorte SOPI do Anala deve continuar na análise e ser classificado como ausente no Anala.

## 9. Contravalidação com o faturamento do mês vigente

Utilize o arquivo de faturamento do mês analisado como terceira fonte de conferência.

No faturamento do mês vigente:

- considere somente operações `1` ou `2`;
- considere somente status `A` ou `E`;
- utilize o código normalizado do PDV;
- utilize o código normalizado do produto;
- mantenha somente SKUs pertencentes à cesta SOPI nacional elegível para o KPI;
- elimine duplicidades de `PDV + SKU`;
- calcule a quantidade de SKUs distintos faturados por PDV;
- registre também a relação dos códigos faturados no mês.

O faturamento do mês vigente não deve ser misturado ao histórico de priorização. Ele deve ser usado para confirmar as compras atuais e impedir que um produto já faturado seja relacionado novamente como não comprado.

## 10. Conciliação entre Ambev, Anala e faturamento

Para cada PDV do universo Ambev, apresente lado a lado:

- Setor;
- Código PDV;
- Nome no Ambev;
- Nome no Anala;
- Dia de Visita do VDE;
- Data da Prox Visita VDE;
- SKUs Ambev;
- SKUs Anala;
- SKUs Faturamento do mês;
- Diferença Anala menos Ambev;
- Diferença Faturamento menos Anala;
- Situação;
- Gap conservador até 10;
- Produtos atuais no Anala, em códigos;
- Produtos faturados no mês, em códigos;
- Dias de entrega disponíveis.

Use uma única classificação de situação por PDV e aplique obrigatoriamente esta precedência, da primeira condição atendida para a última:

1. `Ausente no Anala e no faturamento`: não existe contagem disponível nas duas fontes internas;
2. `Ausente no Anala`: não existe contagem disponível no Anala, mas existe no faturamento;
3. `Ausente no faturamento`: não existe contagem disponível no faturamento, mas existe no Anala;
4. `Convergente 3 fontes — aderente`: as três contagens existem, são exatamente iguais e o valor comum é maior ou igual a 10;
5. `Convergente 3 fontes — abaixo da meta`: as três contagens existem, são exatamente iguais e o valor comum é menor que 10;
6. `Divergente — alguma fonte abaixo de 10`: as três contagens existem, não são iguais e pelo menos uma é menor que 10;
7. `Divergente — todas as fontes em 10 ou mais`: as três contagens existem, todas são maiores ou iguais a 10, mas os números não são exatamente iguais.

Não atribua simultaneamente mais de uma situação ao mesmo PDV. A situação descreve a conciliação; a decisão de ação é derivada separadamente conforme a seção 11.

Calcule:

- `Dif. Anala − Ambev = SKUs Anala − SKUs Ambev`;
- `Dif. Fat. − Anala = SKUs Faturamento − SKUs Anala`.

## 11. Definição do GAP e da ação comercial

Meta mensal: 10 SKUs distintos.

Quando as três contagens estiverem disponíveis, calcule o GAP conservador utilizando as três fontes:

`GAP conservador = máximo entre 0 e (10 − mínimo entre SKUs Ambev, SKUs Anala e SKUs Faturamento do mês)`

Forma equivalente:

`GAP conservador = máximo entre (10 − SKUs Ambev), (10 − SKUs Anala), (10 − SKUs Faturamento do mês) e 0`

Exemplo: Ambev com 4, Anala com 7 e faturamento com 6 resulta em GAP 6.

Quando uma fonte estiver ausente:

- não transforme a ausência em zero;
- calcule um `GAP provisório` usando o menor valor entre as fontes disponíveis;
- mantenha o caso em validação;
- registre explicitamente qual fonte está ausente;
- não considere a meta definitivamente comprovada.

### 11.1. Regra única de decisão

A relação de itens não comprados somente deve ser bloqueada quando:

`SKUs Ambev = SKUs Anala = SKUs Faturamento do mês >= 10`

Essa é a única condição `Sem ação`.

Em todos os demais casos, gere a relação de itens não comprados e classifique a decisão desta forma:

- `Ação comercial`: as três fontes convergem abaixo de 10;
- `Ação comercial + validação`: existe divergência e pelo menos uma fonte está abaixo de 10;
- `Validação com relação preventiva`: todas as fontes disponíveis estão em 10 ou mais, mas os números não coincidem, ou existe alguma fonte ausente sem comprovação completa das três fontes;
- `Fonte ausente + ação comercial e validação`: existe fonte ausente e pelo menos uma das fontes disponíveis está abaixo de 10;
- `Sem ação — 3 fontes convergentes em 10 ou mais`: as três contagens são exatamente iguais e maiores ou iguais a 10.

O fato de apenas uma ou duas fontes registrarem 10 ou mais não bloqueia a relação. Ambev, Anala e faturamento precisam se conversar e apresentar a mesma contagem aderente.

Quando o GAP for zero, mas as fontes forem divergentes ou incompletas, a relação gerada é preventiva e serve para apoiar a validação; não declare automaticamente que o PDV está abaixo da meta.

## 12. Relação completa dos itens da cesta ainda não comprados

Esta etapa é obrigatória em toda reanálise.

Para cada PDV que não esteja classificado como `Sem ação — 3 fontes convergentes em 10 ou mais`:

1. forme o conjunto de todos os SKUs distintos da cesta regional elegível para sugestões;
2. forme o conjunto de SKUs comprados pelo PDV no Anala durante o mês vigente;
3. forme o conjunto de SKUs faturados para o PDV no mês vigente, usando operações 1/2 e status A/E;
4. faça a união dos SKUs encontrados no Anala e no faturamento do mês;
5. mantenha nessa união somente os SKUs pertencentes à cesta regional elegível para sugestões;
6. subtraia essa união do conjunto total da cesta regional elegível para sugestões;
7. relacione todos os SKUs restantes, um por linha;
8. não limite a quantidade de linhas ao GAP;
9. não selecione apenas os itens mais populares;
10. não elimine produtos somente porque não existe histórico do próprio PDV;
11. não inclua produtos da cesta nacional que não tenham sido comercializados pela revenda em 2026;
12. não repita o mesmo SKU para o mesmo PDV.

Regra de fechamento obrigatória por PDV:

`Qtd. itens não comprados = Qtd. SKUs distintos da cesta regional elegível para sugestões − Qtd. SKUs distintos da união Anala + faturamento do mês que pertencem à cesta regional elegível para sugestões`

A união deve contar cada SKU somente uma vez, mesmo que apareça nas duas fontes.

O GAP informa quantos SKUs faltam para chegar a 10 segundo a cesta nacional do KPI. A quantidade de itens não comprados informa todas as opções da cesta regional de sugestões que o PDV ainda não comprou. Esses dois números possuem finalidades diferentes e não devem ser igualados artificialmente.

## 13. Histórico para ordenar os itens

Os relatórios de faturamento de 2026 possuem duas funções diferentes e obrigatórias:

1. todos os meses disponíveis de 2026, inclusive o mês vigente até sua data de corte, definem quais SKUs da cesta nacional são comercializados pela revenda e, portanto, entram na cesta regional elegível para sugestões;
2. somente os meses anteriores ao mês vigente servem como histórico para ordenar os itens da relação de cada PDV.

Depois que a cesta regional elegível para sugestões estiver formada, o histórico serve somente para definir a ordem da relação completa. Ele não deve excluir um item elegível apenas porque o próprio PDV nunca o comprou.

Utilize os arquivos de faturamento anteriores ao mês vigente.

No histórico:

- considere somente operações `1` ou `2`;
- considere somente status `A` ou `E`;
- mantenha somente PDVs do universo SOPI auditado;
- mantenha somente produtos da cesta regional elegível para sugestões;
- considere `Quantidade histórica` como a soma do campo `Qtde`, sem conversão de unidade;
- considere `Valor histórico` como a soma do campo `Valor`;
- calcule quantidade de meses com compra;
- identifique o mês mais recente com compra.

Ordene os itens não comprados de cada PDV nesta sequência:

1. itens com histórico do próprio PDV;
2. itens sem histórico próprio, mas populares no histórico da base SOPI auditada;
3. itens da cesta regional elegível para sugestões sem histórico na base auditada.

Para itens com histórico do próprio PDV, ordenar por:

1. mês mais recente, do mais recente para o mais antigo;
2. maior quantidade de meses com compra;
3. maior quantidade histórica;
4. maior valor histórico;
5. código do SKU como critério final de desempate.

Para itens populares na base SOPI, ordenar por:

1. maior quantidade de PDVs distintos que compraram o SKU;
2. mês mais recente;
3. maior quantidade de meses com compra;
4. maior quantidade histórica;
5. código do SKU como critério final de desempate.

Classifique internamente cada combinação `PDV + SKU não comprado` em exatamente uma categoria, usando esta precedência:

1. `Histórico do próprio PDV`: o próprio PDV comprou o SKU em pelo menos um mês anterior ao mês vigente;
2. `Popular no histórico da base SOPI`: o próprio PDV não possui histórico do SKU, mas pelo menos outro PDV do universo SOPI auditado comprou o produto em mês anterior;
3. `Item da cesta SOPI sem histórico`: não existe compra anterior do SKU nem pelo próprio PDV nem por outro PDV do universo SOPI auditado.

As categorias são mutuamente exclusivas. Um item com histórico próprio nunca deve ser contado novamente como popular. Um item popular nunca deve ser contado novamente como sem histórico.

Essa classificação possui somente duas finalidades:

- ordenar as linhas da aba `Sugestões`;
- alimentar os três totais correspondentes da aba `Base das Sugestões`.

Ela não pode incluir ou excluir produtos da cesta regional, não pode limitar a relação ao GAP e não substitui a decisão de ação definida na seção 11.

## 14. Setor e informações de visita

Cruze os PDVs com o cadastro mais recente da pasta `CLIENTES +`.

Acrescente, em todas as abas que possuem código do PDV:

- `Setor`, imediatamente antes de `Código PDV`;
- `Dia de Visita do VDE`;
- `Data da Prox Visita VDE`.

Para o Setor, priorize o campo oficial do planificador Ambev quando ele representar o código do RN. Utilize o cadastro de clientes como apoio de validação.

Não invente dia ou data de visita. Quando não houver correspondência confiável, mantenha o campo vazio e registre a pendência na validação final.

Grave a próxima visita como data real do Excel, com formato `dd/mm/aaaa`, e não como texto ou número serial sem formatação.

## 15. Dias de entrega

Utilize a fonte mais recente da pasta `DIAS DE ENTREGA`.

Cruze:

- `ERP_POC_ID` com o código normalizado do PDV;
- as colunas de segunda-feira a domingo com os valores `FREE`, `FLEX` ou `NO`.

Na coluna final do relatório, apresente somente os dias disponíveis:

- `SEG - FREE`;
- `SEG - FLEX`;
- `TER - FREE`;
- `TER - FLEX`;
- e assim sucessivamente até domingo.

Separe múltiplos dias por ponto e vírgula.

Não apresente dias com valor `NO`.

Quando o PDV não possuir dia de entrega disponível, mantenha o campo vazio.

## 15.1. Preço bairro

Utilize a fonte mais recente da pasta `TABELA DE PREÇOS`.

Na fonte de preços:

- utilize a primeira coluna como código do SKU;
- utilize a coluna J, identificada pelo cabeçalho `VALOR`, como preço de venda do produto;
- normalize o código do SKU com as mesmas regras aplicadas à cesta e às compras;
- trate o código como texto para o cruzamento;
- trate `VALOR` como número monetário, e não como texto formatado;
- elimine repetições quando o mesmo SKU possuir exatamente o mesmo preço em todas as ocorrências;
- não escolha silenciosamente um valor quando o mesmo SKU possuir preços diferentes: mantenha `Preço bairro` vazio para esse SKU, registre o conflito e informe a pendência na resposta final;
- não transforme SKU ausente, preço vazio ou preço inválido em zero;
- não procure preço por nome ou descrição do produto quando o código do SKU não corresponder.

Na aba `Sugestões`:

1. cruze `SKU não comprado` com o código normalizado da primeira coluna da tabela de preços;
2. grave o valor correspondente na coluna `Preço bairro`;
3. mantenha a célula vazia quando o SKU não existir na fonte, o valor estiver inválido ou houver conflito entre preços;
4. formate a coluna como moeda brasileira, preservando o valor como número real do Excel;
5. posicione a coluna imediatamente depois de `Produto da cesta não comprado` e antes de `Dias de entrega disponíveis`;
6. não acrescente a coluna de preço às abas `Conciliação` ou `Base das Sugestões`;
7. não altere a quantidade, a seleção ou a ordem das sugestões por causa do preço.

O preço é uma informação comercial complementar. Ele não participa do cálculo do KPI, do GAP, da cesta regional, da classificação, da priorização histórica ou da definição de ação.

## 16. Estrutura obrigatória do Excel

O relatório deve possuir exatamente estas três abas, nesta ordem:

1. `Sugestões`;
2. `Conciliação`;
3. `Base das Sugestões`.

Não crie abas adicionais sem autorização do usuário.

Em especial, não crie uma aba chamada `Resumo`, `SKUPDV - CESTA REGIONAL`, `Cesta Regional`, `Cesta SOPI Regional` ou equivalente. A aplicação da cesta regional deve ocorrer dentro da aba `Sugestões`, sem mudar a estrutura de três abas.

Os totais gerais, a metodologia, as fontes utilizadas, as datas de corte e as pendências devem ser informados na resposta final, e não em uma aba adicional.

### 16.1. Aba Sugestões

Apesar do nome padronizado da aba, ela deve conter a relação completa dos itens da cesta regional elegível para sugestões ainda não comprados, e não apenas uma quantidade de sugestões igual ao GAP.

Inclua nessa aba todos os PDVs que não estejam classificados como `Sem ação — 3 fontes convergentes em 10 ou mais`, inclusive casos de validação com relação preventiva e GAP zero.

Utilize exatamente estas colunas, nesta ordem:

1. `Setor`;
2. `Código PDV`;
3. `Nome PDV`;
4. `Dia de Visita do VDE`;
5. `Data da Prox Visita VDE`;
6. `GAP`;
7. `SKU não comprado`;
8. `Produto da cesta não comprado`;
9. `Preço bairro`;
10. `Dias de entrega disponíveis`.

Cada relação `PDV + SKU` deve ocupar uma linha.

Repita os dados do PDV em todas as linhas para permitir filtro, pesquisa e exportação.

Não apresente somente o código do produto. Apresente código e descrição.

### 16.2. Aba Conciliação

Utilize exatamente estas colunas, nesta ordem:

1. `Setor`;
2. `Código PDV`;
3. `Nome no Ambev`;
4. `Nome no Anala`;
5. `Dia de Visita do VDE`;
6. `Data da Prox Visita VDE`;
7. `SKUs Ambev`;
8. `SKUs Anala`;
9. `SKUs Faturamento do mês`;
10. `Dif. Anala − Ambev`;
11. `Dif. Fat. − Anala`;
12. `Situação`;
13. `Gap conservador até 10`;
14. `Produtos atuais no Anala (códigos)`;
15. `Produtos faturados no mês (códigos)`;
16. `Dias de entrega disponíveis`.

Cada PDV auditado deve ocupar exatamente uma linha.

### 16.3. Aba Base das Sugestões

Para evitar duplicar centenas de milhares de linhas, esta aba deve resumir a composição da relação completa por PDV.

Cada PDV do universo auditado deve ocupar exatamente uma linha, inclusive os PDVs sem ação.

Utilize exatamente estas colunas, nesta ordem:

1. `Setor`;
2. `Código PDV`;
3. `Nome PDV`;
4. `Dia de Visita do VDE`;
5. `Data da Prox Visita VDE`;
6. `SKUs Ambev`;
7. `SKUs Anala`;
8. `SKUs Faturamento do mês`;
9. `Gap`;
10. `Qtd. itens não comprados`;
11. `Com histórico próprio`;
12. `Populares na base SOPI`;
13. `Sem histórico`;
14. `Mês histórico mais recente`;
15. `Quantidade histórica própria`;
16. `Valor histórico próprio`;
17. `Cobertura da cesta`;
18. `Classificação`;
19. `Dias de entrega disponíveis`.

Regra de fechamento da composição:

`Com histórico próprio + Populares na base SOPI + Sem histórico = Qtd. itens não comprados`

Calcule a cobertura da cesta como:

`SKUs distintos comprados na união Anala + faturamento que pertencem à cesta regional elegível para sugestões ÷ SKUs distintos da cesta regional elegível para sugestões`

Essa cobertura é um indicador comercial da relação de sugestões. Ela não substitui a aderência oficial do KPI, que continua sendo calculada pela meta de 10 SKUs distintos da cesta SOPI nacional.

Definições obrigatórias dos campos desta aba:

- `SKUs Ambev`: contagem distinta informada na fonte Ambev, sem substituição por outra fonte;
- `SKUs Anala`: contagem distinta recalculada com os produtos do Anala pertencentes à cesta nacional vigente;
- `SKUs Faturamento do mês`: contagem distinta recalculada no faturamento vigente com operação 1/2, status A/E e cesta nacional;
- `Gap`: GAP conservador ou provisório definido na seção 11;
- `Qtd. itens não comprados`: quantidade de linhas distintas `PDV + SKU` geradas para o PDV na aba `Sugestões`;
- `Com histórico próprio`: quantidade de SKUs sugeridos classificados como `Histórico do próprio PDV`;
- `Populares na base SOPI`: quantidade de SKUs sugeridos classificados como `Popular no histórico da base SOPI`;
- `Sem histórico`: quantidade de SKUs sugeridos classificados como `Item da cesta SOPI sem histórico`;
- `Mês histórico mais recente`: mês mais recente, anterior ao mês vigente, em que o PDV comprou qualquer um dos SKUs atualmente relacionados como não comprados e classificados com histórico próprio; deixar vazio quando não existir histórico próprio;
- `Quantidade histórica própria`: soma do campo `Qtde`, sem conversão de unidade, somente para o próprio PDV, somente nos meses anteriores ao vigente e somente para os SKUs atualmente relacionados que possuem histórico próprio;
- `Valor histórico próprio`: soma do campo `Valor` sob o mesmo universo da quantidade histórica própria;
- `Cobertura da cesta`: quantidade de SKUs distintos da cesta regional encontrados na união Anala + faturamento vigente, dividida pela quantidade total de SKUs distintos da cesta regional;
- `Classificação`: decisão final da seção 11, utilizando exclusivamente um dos rótulos `Ação comercial`, `Ação comercial + validação`, `Validação com relação preventiva`, `Fonte ausente + ação comercial e validação` ou `Sem ação — 3 fontes convergentes em 10 ou mais`;
- `Dias de entrega disponíveis`: resultado do cruzamento definido na seção 15.

Para PDVs classificados como `Sem ação`, grave `Qtd. itens não comprados`, `Com histórico próprio`, `Populares na base SOPI` e `Sem histórico` como zero; mantenha os campos históricos vazios quando não houver relação de itens a resumir.

## 17. Layout e formatação padronizados

Utilize como referência obrigatória de layout o arquivo:

`C:\Users\matheus.manoel\Downloads\excel-modelo.xlsx`

O objetivo é produzir um relatório operacional simples, funcional, filtrável e fácil de exportar. Não reproduza o antigo layout com faixas azuis, cartões, títulos decorativos ou blocos separados.

### 17.1. Estrutura das abas

Aplicar estas regras nas três abas:

- nenhuma célula mesclada;
- nenhuma linha decorativa ou área vazia antes dos dados;
- o intervalo utilizado deve começar obrigatoriamente em `A1`;
- o cabeçalho deve ocupar a primeira linha;
- o intervalo deve ocupar somente as linhas e colunas necessárias;
- não formatar linhas ou colunas vazias fora do intervalo utilizado;
- não ocultar linhas ou colunas obrigatórias;
- não incluir gráficos, imagens, formas, cartões, caixas de texto, ícones decorativos ou outros objetos;
- não criar títulos acima da tabela;
- informações de título, período, metodologia, fontes, totais gerais e pendências devem ser apresentadas na resposta final, sem criação de uma aba adicional.

### 17.2. Tabelas estruturadas

Transforme o intervalo preenchido de cada aba em uma Tabela estruturada do Excel.

Use nomes únicos e válidos:

- aba `Sugestões`: `tblSugestoes`;
- aba `Conciliação`: `tblConciliacao`;
- aba `Base das Sugestões`: `tblBaseSugestoes`.

Configuração obrigatória das tabelas:

- estilo `TableStyleLight1`;
- cabeçalho visível;
- filtros ativos em todas as colunas;
- linha de totais desativada;
- linhas alternadas ativadas;
- colunas alternadas desativadas;
- destaque da primeira coluna desativado;
- destaque da última coluna desativado;
- uma única tabela por aba;
- nenhuma tabela sobreposta ou com intervalo maior que a área preenchida.

Se a quantidade de relações ultrapassar o limite físico de linhas de uma planilha do Excel, não reduza, corte, selecione amostras nem substitua silenciosamente a tabela por um intervalo comum. Interrompa a geração, informe o volume encontrado e solicite autorização para dividir a saída.

### 17.3. Fonte, cabeçalhos e linhas

Aplicar em todas as abas:

- fonte `Aptos Narrow`;
- tamanho `11`;
- cabeçalho em negrito;
- cabeçalho com texto preto;
- corpo com fonte normal e texto preto;
- cores, faixas e separações fornecidas pelo estilo nativo `TableStyleLight1`;
- linhas alternadas em branco/cinza muito claro e cinza-claro;
- altura padrão das linhas de aproximadamente `14,4` pontos;
- mesma altura padrão no cabeçalho e no corpo, salvo impossibilidade real de leitura;
- quebra automática de texto desativada;
- textos em uma única linha;
- alinhamento horizontal `Geral`: textos à esquerda e números à direita;
- alinhamento vertical inferior, seguindo o modelo;
- linhas de grade visíveis.

Não aplique preenchimentos manuais para imitar as faixas da tabela. Não aplique bordas pesadas ou bordas decorativas em todas as células. Utilize somente a separação discreta do `TableStyleLight1` e as linhas de grade.

### 17.4. Largura das colunas

Ajuste as larguras ao conteúdo, mantendo as linhas em uma única altura e evitando espaços excessivos.

Use como referência para a aba `Sugestões`:

- `Setor`: aproximadamente `7,11`;
- `Código PDV`: aproximadamente `12,22`;
- `Nome PDV`: aproximadamente `22,11`;
- `Dia de Visita do VDE`: aproximadamente `18,66`;
- `Data da Prox Visita VDE`: aproximadamente `21,33`;
- `GAP`: aproximadamente `6,44`;
- `SKU não comprado`: aproximadamente `18,00`;
- `Produto da cesta não comprado`: aproximadamente `62,55`;
- `Preço bairro`: aproximadamente `16,00`;
- `Dias de entrega disponíveis`: aproximadamente `58,44`.

Nas demais abas:

- colunas de códigos e contagens devem ser compactas;
- colunas de nomes devem ser largas o suficiente para leitura;
- colunas de descrição, produtos, situação, classificação, observações, fontes e dias de entrega devem ser mais largas;
- limite larguras excessivas que prejudiquem a navegação;
- não use quebra automática como substituto para uma largura de coluna adequada.

### 17.5. Formatos de dados obrigatórios

- `Código PDV`, códigos de produtos, `SKU não comprado`, listas de SKUs e outros identificadores devem ser armazenados como texto;
- nunca permita notação científica para PDV ou SKU;
- preserve zeros à esquerda no valor exibido quando existirem na fonte;
- `Data da Prox Visita VDE` deve ser armazenada como data real do Excel e exibida em `dd/mm/aaaa`;
- não deixe datas no formato `Geral`, evitando números seriais como `46262`;
- contagens e GAPs devem usar formato inteiro, sem casas decimais;
- valores históricos devem usar formato monetário adequado;
- `Preço bairro` deve ser armazenado como número e exibido em formato monetário, com duas casas decimais;
- coberturas e percentuais devem usar formato percentual;
- campos vazios devem permanecer vazios, sem zeros artificiais quando zero não representar um valor real.

### 17.6. Painéis, filtros e usabilidade

- mantenha os filtros nativos das tabelas;
- não utilize segmentações, botões ou controles adicionais;
- não congele quatro linhas nem aplique o antigo congelamento das colunas `Setor` e `Código PDV`;
- para copiar fielmente o modelo, mantenha os painéis sem congelamento;
- permita pesquisa, ordenação, filtragem, cópia e exportação direta das linhas;
- preserve a ordem das três abas e a ordem das colunas definidas na seção 16.

## 18. Caminho e nome do relatório

Salve o relatório obrigatoriamente dentro de:

`C:\Users\matheus.manoel\OneDrive - Vyas\Matheus - NS - Sinop\SPO - Gestão A Vista\SKU-PDV\SKUPDV - <mês por extenso>\DD-MM-AAAA`

Exemplo para agosto:

`C:\Users\matheus.manoel\OneDrive - Vyas\Matheus - NS - Sinop\SPO - Gestão A Vista\SKU-PDV\SKUPDV - agosto\24-08-2026`

Crie a pasta do mês e a pasta da data apenas quando ainda não existirem.

Nome padrão:

`Relatorio_KPI_SKU_PDV_SOPI_AAAA-MM-DD.xlsx`

Nunca sobrescreva um relatório anterior.

Se o nome já existir, utilize:

- `_v2`;
- `_v3`;
- `_v4`;
- e assim sucessivamente.

Não salve o relatório fora do caminho oficial.

## 19. Validações numéricas obrigatórias

Antes da entrega, confirme:

1. quantidade de PDVs do planificador Ambev;
2. quantidade de linhas da aba `Conciliação` igual à quantidade de PDVs auditados;
3. ausência de PDVs duplicados na `Conciliação`;
4. todos os setores preenchidos ou pendências identificadas;
5. todos os SKUs relacionados pertencentes simultaneamente à cesta SOPI nacional e à cesta regional elegível para sugestões;
6. todos os SKUs da cesta regional de sugestões possuem pelo menos um faturamento elegível da revenda em 2026;
7. nenhum SKU da cesta nacional sem faturamento elegível da revenda em 2026 aparece na aba `Sugestões`;
8. nenhum SKU já encontrado no Anala relacionado novamente para o mesmo PDV;
9. nenhum SKU já encontrado no faturamento elegível do mês relacionado novamente para o mesmo PDV;
10. nenhuma duplicidade de `PDV + SKU` na aba `Sugestões`;
11. quantidade de itens por PDV igual à cesta regional elegível para sugestões menos a união das compras atuais pertencentes a essa cesta regional;
12. soma das três classificações históricas igual à quantidade de itens não comprados;
13. ausência de relação comercial somente nos PDVs em que Ambev, Anala e faturamento possuem contagens disponíveis, exatamente iguais e maiores ou iguais a 10;
14. presença da relação preventiva quando as três fontes estiverem em 10 ou mais, mas com contagens diferentes;
15. presença de ação comercial ou ação com validação quando pelo menos uma fonte disponível estiver abaixo de 10;
16. contagens do KPI em Ambev, Anala e faturamento mantidas sobre a cesta SOPI nacional, sem substituição pela cesta regional de sugestões;
17. total de convergentes + divergentes + ausentes compatível com o universo auditado;
18. códigos tratados como texto, sem perda de zeros ou conversão para notação científica;
19. datas armazenadas como datas reais do Excel;
20. fórmulas sem `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?` ou `#N/A`;
21. caminhos das fontes registrados corretamente;
22. todos os meses de faturamento de 2026 utilizados na cesta regional estão identificados e eventuais lacunas estão registradas;
23. relatório anterior preservado;
24. quantidade de abas igual a três, na ordem definida neste prompt;
25. inexistência de aba separada para a cesta regional;
26. colunas `SKU não comprado` e `Produto da cesta não comprado` contendo somente o resultado da cesta regional menos as compras atuais do PDV;
27. fechamento das categorias históricas sem sobreposição;
28. GAP calculado com as três fontes quando disponíveis e marcado como provisório quando houver ausência;
29. uma Tabela estruturada por aba, iniciada em `A1` e limitada ao intervalo preenchido;
30. nomes das tabelas iguais a `tblSugestoes`, `tblConciliacao` e `tblBaseSugestoes`;
31. estilo `TableStyleLight1`, filtros ativos, linhas alternadas e linha de totais desativada;
32. inexistência de células mescladas nas três abas;
33. inexistência de gráficos, imagens, formas, cartões ou caixas de texto;
34. códigos de PDV e SKU armazenados como texto;
35. datas de próxima visita armazenadas como datas reais e exibidas em `dd/mm/aaaa`;
36. coluna `Preço bairro` presente somente na aba `Sugestões`, imediatamente após `Produto da cesta não comprado`;
37. preços cruzados exclusivamente pelo código normalizado do SKU com a primeira coluna da fonte de preços;
38. valores obtidos exclusivamente da coluna J `VALOR`, armazenados como números e formatados como moeda;
39. SKUs sem preço válido mantidos em branco, sem conversão artificial para zero;
40. duplicidades com preços conflitantes identificadas e registradas, sem escolha silenciosa de um valor;
41. conteúdo, fórmulas, estrutura e formatação das abas `Conciliação` e `Base das Sugestões` preservados quando a única solicitação nova for a inclusão do preço.

Faça testes por amostragem em PDVs com:

- GAP 1;
- GAP alto;
- nenhuma compra no mês;
- compra presente no Anala e no faturamento;
- divergência entre Anala e Ambev;
- ausência no Anala;
- múltiplos dias de entrega;
- histórico próprio;
- somente popularidade da base;
- itens sem histórico.

## 20. Validação visual obrigatória

Antes de entregar:

- abra ou renderize todas as abas;
- confirme que todas as tabelas começam em `A1`;
- confirme que não existem títulos decorativos ou linhas vazias antes do cabeçalho;
- confira se todos os cabeçalhos estão visíveis, em Aptos Narrow 11, negrito e texto preto;
- confirme o estilo `TableStyleLight1` e a alternância das linhas;
- confirme que não existe nenhuma célula mesclada;
- confirme que não existem gráficos, imagens, formas ou cartões;
- confira se não existem textos sobrepostos;
- confira se datas não aparecem como números seriais;
- confira se códigos não aparecem em notação científica;
- confira se códigos de PDV e SKU foram preservados como texto;
- confira se nomes e produtos não estão cortados de forma ilegível;
- confira se as cores alternadas permanecem consistentes;
- confira se os filtros estão ativos em todas as tabelas;
- confira se não há linha de totais;
- confira se não há congelamentos herdados do layout antigo;
- confira se a quebra automática está desativada e as linhas mantêm altura próxima de 14,4 pontos;
- confira se não existem colunas obrigatórias ocultas;
- confira se `Preço bairro` está entre a descrição do produto e os dias de entrega, com valores monetários legíveis e células vazias para SKUs sem preço válido;
- confira as primeiras e as últimas linhas das abas extensas;
- confirme que o arquivo abre corretamente no Excel.

## 21. Resposta final

Na resposta final:

- entregue o link do novo relatório;
- informe o caminho em que ele foi salvo;
- confirme que os relatórios anteriores foram preservados;
- informe a data das fontes utilizadas;
- informe a quantidade de PDVs auditados;
- informe convergentes, divergentes e ausentes;
- informe PDVs com ação comercial;
- informe PDVs com ação comercial + validação;
- informe PDVs com validação e relação preventiva;
- informe PDVs sem ação por convergência exata das três fontes em 10 ou mais;
- informe casos para validação de sistemas;
- informe a quantidade de SKUs distintos da cesta SOPI nacional usada no KPI;
- informe a quantidade de SKUs distintos da cesta regional elegível para sugestões;
- informe quantos SKUs da cesta nacional foram excluídos das sugestões por não terem faturamento elegível da revenda em 2026;
- informe a quantidade total de relações `PDV + SKU ainda não comprado`;
- informe eventuais pendências de setor, visita, entrega ou diferença de corte;
- informe a fonte e a data da tabela de preços, a quantidade de SKUs com preço, a quantidade de SKUs sugeridos sem preço e eventuais conflitos de preço;
- não apresente como completa uma lista que represente somente um subconjunto do universo analisado.

---
