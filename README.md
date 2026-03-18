# Protótipo VBA para Automatização de Cálculos em Metrologia Elétrica

![VBA](https://img.shields.io/badge/VBA-Excel-blue)
![Status](https://img.shields.io/badge/status-prototype-orange)
![Domain](https://img.shields.io/badge/domain-Metrology-green)
![Quality](https://img.shields.io/badge/focus-Measurement%20Reliability-important)

---
## Sobre o projeto
Este projeto consiste no desenvolvimento de um protótipo em Visual Basic for Applications (VBA) voltado para a automatização de cálculos em planilhas utilizadas em metrologia elétrica.

A solução foi concebida para substituir processos manuais em planilhas, reduzindo erros operacionais e aumentando a confiabilidade no tratamento de dados de medição.

O algoritmo implementado realiza busca automatizada em tabelas técnicas, considerando diferentes escalas de unidades elétricas, além de garantir a consistência dimensional dos resultados.
## Valor da Solução

-  Redução de erros em cálculos manuais  
-  Aumento da confiabilidade nos processos de calibração  
-  Padronização do tratamento de dados de medição  
-  Melhoria na rastreabilidade dos cálculos  
-  Integração direta com fluxos existentes baseados em Excel

 ## Principais Funcionalidades

-  Busca automatizada em tabelas de referência  
-  Conversão entre múltiplas escalas (µ, m, k, M, G, T)  
-  Normalização de unidades para comparação consistente  
-  Retorno de resultados ajustados à unidade de entrada  
-  Identificação de valores fora da faixa ("Fora de Range")

##  Abordagem Técnica

O algoritmo foi estruturado em quatro etapas principais:

1. **Normalização da entrada**  
   Converte o valor informado para uma base comum utilizando o multiplicador da unidade.

2. **Varredura da tabela**  
   Percorre as linhas da tabela, convertendo os limites para a mesma base de comparação.

3. **Identificação da faixa**  
   Determina a faixa de medição aplicável com base na comparação dos valores.

4. **Ajuste do resultado**  
   Converte o valor retornado para a mesma escala da unidade de entrada.

Essa abordagem garante consistência independentemente da ordem de grandeza das unidades envolvidas.

---

##  Abordagem Técnica

O algoritmo foi estruturado em quatro etapas principais:

1. **Normalização da entrada**  
   Converte o valor informado para uma base comum utilizando o multiplicador da unidade.

2. **Varredura da tabela**  
   Percorre as linhas da tabela, convertendo os limites para a mesma base de comparação.

3. **Identificação da faixa**  
   Determina a faixa de medição aplicável com base na comparação dos valores.

4. **Ajuste do resultado**  
   Converte o valor retornado para a mesma escala da unidade de entrada.

Essa abordagem garante consistência independentemente da ordem de grandeza das unidades envolvidas.


##  Arquitetura do Código

### ➔ Função principal
`BuscarEletricoUnificado`

Responsável por:
- Busca na tabela  
- Normalização das unidades  
- Identificação da faixa  
- Ajuste do resultado  

### ➔ Função auxiliar
`ObterMultiplicador`

Responsável por:
- Identificar o prefixo métrico  
- Retornar o fator de multiplicação correspondente (10ⁿ)  


## Aplicações

Este projeto pode ser aplicado em:

- Laboratórios de calibração  
- Processamento de dados de medição  
- Automação de planilhas técnicas  
- Análise de resultados experimentais  
- Sistemas de controle de qualidade  


## Limitações

- Não possui interface gráfica dedicada  
- Validação de entradas ainda limitada  
- Dependência da estrutura correta das tabelas no Excel  
- Protótipo ainda não validado em ambiente operacional completo  


## Melhorias Futuras

- Implementação de interface gráfica (UserForm)  
- Validação robusta de entradas e tratamento de exceções  
- Expansão para outras grandezas físicas  
- Implementação de testes automatizados  
- Alinhamento completo com ISO/IEC 17025  

---

## Tecnologias Utilizadas

- Microsoft Excel  
- Visual Basic for Applications (VBA)  

 Como Utilizar

1. Abra o Microsoft Excel  
2. Pressione `ALT + F11` para acessar o Editor VBA  
3. Importe o módulo `.bas` para o projeto  
4. Utilize a função diretamente nas células  

### Exemplo:

```excel
=BuscarEletricoUnificado(A1;"kO";Tabela;1;2;3)
