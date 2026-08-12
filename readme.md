# Processamento dos dados de Investimentos Públicos no Estado do Ceará


## Propósito

<p>

      Este repositório[^1] contém as rotinas responsáveis por realizar
os devidos tratamentos nos dados de investimentos públicos no Estado do
Ceará necessários ao projeto. Em termos práticos, os dados de
investimentos são de frequência mensal e de caráter cumulativo. Dada
essa estrutura, se faz necessário tomar a variação entre meses de modo a
obter o real investimento no período.

</p>

## Sobre os dados

<p>

      Os dados são obtidos por meio de planilhas geradas no Sistema
Integrado Orçamentário e Financeiro do Estado do Ceará
[SIOF](https://planejamento.seplag.ce.gov.br/siofconsulta/Paginas/frm_consulta_execucao.aspx)
disponibilizados pela Secretaria do Planejamento e Gestão do Estado do
Ceará [SEPLAG](https://www.ce.gov.br/seplag/). Os dados coletados são de
frequência mensal e são divididos entre **Investimentos por Programa** e
**Investimentos por Função**. Além disso, para o primeiro grupo de
informações os investimentos são classificados pelos seguintes tipos:
**Equipamentos**, **Obras** e **Total**. No grupo dos **Investimentos
por Função**, juntam-se aos tipos mencionados os Investimentos
**Correntes** e com **Pessoal**.  
      Como mencionado, os dados possuem frequência mensal e são de
caráter cumulativo. Portanto, para se chegar ao investimento efetivo de
cada período é necessário tomar a diferença entre cada um dos intervalos
seguindo a seguinte regra[^2]:

</p>

$$invest_t = invest\_acum_t - invest\_acum_{t-1}    \qquad(t \neq 1)$$

<p>

      A base de dados de Investimentos por Programa está também
identificada por região. Dessa maneira, existe a possibilidade de
acessar a informação de três maneiras distintas: i) Programa, ii) Região
e iii) Programa e Região; atendendo as necessidades do projeto.

      Com respeito ao **time span** do conjunto de dados, a série de
**Investimentos por Programa** inicia-se a partir de 2016, em razão dos
período 2013-2015 não ter disponível as informações por região. Para a
série de **Investimentos por Função**, a série tem início em 2015.

</p>

## Rotinas

      As rotinas de tratamento foram desenvolvidas na linguagem Python e
estão detalhadas no link abaixo:

- [**data_processing_investimentos_funcao**](https://github.com/paulo-icaro/Investimentos_Publicos_Sefaz/blob/main/data_processing_investimentos_funcao.md):
  investimentos por função;
- [**data_processing_investimentos_programa_regiao**](https://github.com/paulo-icaro/Investimentos_Publicos_Sefaz/blob/main/data_processing_investimentos_programa_regiao.md):
  investimentos por programa e região.

[^1]: É importante ressaltar ambas rotinas foram geradas usando a
    interface [Spyder](https://www.spyder-ide.org/) que possui um
    *spyproject* associado que facilita a execução e leitura dos
    arquivos contidos em planilha. Em razão do tamanho dessas planilhas,
    tais dados não estão disponibilizados neste repositório. Contudo,
    podem ser acessados na página do
    [SIOF](https://planejamento.seplag.ce.gov.br/siofconsulta/Paginas/frm_consulta_execucao.aspx)
    ou via solicitação a este autor. Os demais arquivos são referente a
    geração dessa documentação e não são pertinentes a execução das
    análises do projeto.

[^2]: Obviamente, para o mês de janeiro não é necessário realizar nenhum
    tipo de transformação
