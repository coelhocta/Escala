<p align="center"><img src="https://github.com/coelhocta/Escala/blob/master/img/soldado.ico"></p>

# Gerador de Escala de Serviço

## Pré-requisitos

* Windows.
* Microsoft Office Excel 2007.

## Instruções:

* Abra a planilha **Escala.xlsx** e selecione a aba **Inicio**:

* **PERÍODO:** Escolha o período que a escala será gerada. Célula **B1** Início e **C1** término.

* **Casos especiais:** Se um feriado cai durante a semana, deve ser inserido o(s) dia(s) do feriado ao lado da cor correspondente.

  - ***1:*** Se o Feriado cai numa sexta-feira, deve-se inserir a data do feriado ao lado da célula ***Vermelha***, o sistema automaticamente vai colocar um dia antes como escala marrom.

  - ***2:*** Se o feriado cai numa quarta feira e na terça deve ser marrom, deve se inserir a data do feriado ao lado da célula ***Vermelha***, e deve-se inserir a data da escala marrom ao lado da célula ***Marrom***.

* **Assinaturas:** Nas células ***F1*** e ***J1*** devem ser inseridos as assinaturas e nas células ***F2*** e ***J2*** devem ser inseridas as Funções.

* **Militares que concorrem à Escala:** Na ***Coluna A*** a partir da célula ***8*** devem ser inseridos os nomes de todos os Militares **POR ANTIGUIDADE** que concorrem à escala de Serviço e **DEVE-SE INSERIR EM TODAS AS ABAS, Vermelha, Preta, Marrom e Roxa!**

* **INDISPONIBILIDADES:** Insira todos os dias indisponíveis ao lado de cada militar indisponível.

## Criação da Escala:

1. Após ter preenchido as planilhas, feche a planilha:
2. Execute o programa ***Escala.exe***
3. Será gerado outras planilha completa e preenchida.

## Possíveis erros ou inconsistências

* Usar um editor de planilhas que não seja o Excel 2007.
* Colocar uma data que não existe, por exemplo 31/06/2020.
* Colocar uma data final menor que a data inicial.
* Preencher um quadrinho que não seja correspondente à cor. ***Obs.*** Você pode forçar um militar a tirar serviço no dia que vc quiser, porém preencha com o dia correto ou será gerado uma escala inconsistente com os quadrinhos. *Ex.* Se no dia 19/06/2020 (marrom) eu quero escalar uma pessoa, e escrevo esta data na aba *Vermelha*, isso o sistema não pode tratar, pois o quadrinho deveria estar escrito na aba *Marrom* e não na aba *Vermelha*.
* Não fechar a planilha antes de executar o programa *Escala.exe*

**Observação:** Todas as datas devem ser inseridas no formato **DD/MM/AAAA**, *Exemplo: 01/01/2021*.


## For DEVs:
```sh
pip install -r requirements.txt
para converter py em exe, execute: `auto-py-to-exe`
```

