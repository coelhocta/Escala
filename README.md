<p align="center"><img src="https://github.com/coelhocta/Escala/blob/master/img/soldado.ico"></p>

# Gerador de Escala de Serviço

## Pré-requisitos

* Windows.
* Microsoft Office Excel 2007.

## Instruções

* Abra a planilha **Escala.xlsx** e selecione a aba **Inicio**:

* **PERÍODO:** Escolha o período que a escala será gerada. Célula **B1** Início e **C1** término.

* **Casos especiais:** Se um feriado cai durante a semana, deve ser inserido o(s) dia(s) do feriado ao lado da cor correspondente.

  - ***1:*** Se o Feriado cai numa sexta-feira, deve-se inserir a data do feriado ao lado da célula ***Vermelha***, o sistema automaticamente vai colocar um dia antes como escala marrom.

  - ***2:*** Se o feriado cai numa quarta feira e na terça deve ser marrom, deve se inserir a data do feriado ao lado da célula ***Vermelha***, e deve-se inserir a data da escala marrom ao lado da célula ***Marrom***.

* **Assinaturas:** Nas células ***F1*** e ***J1*** devem ser inseridos as assinaturas e nas células ***F2*** e ***J2*** devem ser inseridas as Funções.

* **Militares que concorrem à Escala:** Na ***Coluna A*** a partir da ***Linha 8*** devem ser inseridos os nomes de todos os Militares **POR ANTIGUIDADE** que concorrem à escala de Serviço e **DEVE-SE INSERIR EM TODAS AS ABAS, Vermelha, Preta, Marrom e Roxa! (Devem ser exatamente Iguais)**

* **INDISPONIBILIDADES:** Insira todos os dias indisponíveis ao lado de cada militar indisponível.

* **Escolher um dia específico para um militar tirar o serviço:** Para escolher um dia para o militar tirar o serviço, basta escolher a aba correspondente à cor da escala, e colocar o dia específico no lugar do quadrinho correspondente, acrescentando um quadrinho para o militar.

* **Contar quadrinhos Premios ou Lastro:** Para contar um quadrinho sem um dia para o militar, basta colocar a palavra ***Lastro*** dentro do quadrinho ao lado do nome dele, como se fosse um quadrinho com a data correta.

* **Para que uma pessoa não concorra à escala Preta ou Marrom:** Escolha a aba correspondente ao quadrinho a qual o militar não deve tirar o serviço e acrescente um **\*** no início do nome, exemplo, se o militar ***"2S SIN Fulano"*** não pode concorrer à escala **Preta e Marrom** então vá na aba ***Preta*** e no lugar do nome, deve ficar desta forma: ***"\* 2S SIN Fulano"***, faça isso também na aba ***Marrom***.

* ***IMPORTANTE: ⭐️ SE O MILITAR NÃO CONCORRE À ESCALA PRETA OU MARROM, MESMO ASSIM O NOME DELE DEVE ESTAR NA CONTAGEM DOS QUADRINHOS PRETA E MARROM.⭐️***

## Criação da Escala

1. Após ter preenchido as planilhas, feche a planilha:
2. Execute o programa ***Escala.exe***
3. Será gerado outras planilha completa e preenchida.

## Possíveis erros ou inconsistências

* Se utilizar um editor de planilhas que não seja o Excel 2007 o sistema não funcionará.
* Colocar uma data que não existe, por exemplo 31/06/2020, dará erro, pois o excel pode considerar o ano errado.
* Colocar uma data final menor que a data inicial, aparecerá erro, a data final sempre deverá ser maior que a data inicial.
* Preencher um quadrinho que não seja correspondente à cor. ***Obs.*** Você pode forçar um militar a tirar serviço no dia que vc quiser, porém preencha com o dia correto ou será gerado uma escala inconsistente com os quadrinhos. *Ex.* Se no dia 19/06/2020 (marrom) eu quero escalar uma pessoa, e escrevo esta data na aba *Vermelha*, isso o sistema não pode tratar, pois o quadrinho deveria estar escrito na aba *Marrom* e não na aba *Vermelha*.
* Se não fechar a planilha antes de executar o programa *Escala.exe* o programa tentará substituir o arquivo de escala e não conseguirá, pois estaria aberto, é necessário que a planilha do excel esteja fechada antes de executar o programa *Escala.exe*.

**Observação:** Todas as datas devem ser inseridas no formato **DD/MM/AAAA**, *Exemplo: 01/01/2021*.

## For DEVs

```sh
pip install -r requirements.txt
para converter py em exe, execute: `auto-py-to-exe`
```
