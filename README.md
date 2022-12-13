# dda-lib-ads-Ativos-Automacao-DF
Automação para preparação do formulário A4 em auditorias de Fundos de Investimentos.


Automação em macro do excel (VBA) para automatizar preenchimento do papel de trabalho formulário A4 da equipe de auditoria.

# Acelerador: Macro para automação de DFs (VBA)


Atualizado por: Vassilis Ioannis Konsolakis
Última Versão: 13-12-2022

## 1-	Descrição
Macro desenvolvida para automatizar os processos de demonstrativos financeiros para FSO. 

**Lista de procedimentos desenvolvidos no VBA:**
1) Leitura das pastas contendo todos os fundos que serão processados;
2) Leitura da planilha excel "Planilha Informações - Automação DF.xlsx" para capturar todas as informações que serão incluídas no modelo;
3) Consolidação dos arquivos das pastas referentes aos fundos processados ao modelo principal;
4) Substituição dos campos "XXXn" no modelo principal de acordo com o fundo processado;
5) Verificação dos valores de PAAs para inclusão do modelo PAA ao modelo principal;

## 2-	Pré Requisitos

Para executar a automação precisa montar no diretório: “C:\” uma pasta chamada: “AutomacaoDF” com os arquivos: Modelo 1.docx, Modelo 2.docx, (SALVAR SEMPRE ABERTO NA SHEET 1) Planilha Informações - Automação DF.xlsx e criar quatro pastas chamadas:

•	Fundos
•	Log
•	PAAs

Adicionar os fundos das administradoras em suas respectivas pastas, criadas anteriormente. Para cada fundo, o nome da pasta deve conter o código de 5 mais “_N”, para DF de data base, ou “_E” para DF de Evento (exemplo: Fundo: 41090, para 41090_N; Fundo: 50037, para: 50037_E).

Para os arquivos, verificar se a planilha “Planilha Informações - Automação DF.xlsx”  está preenchida. As colunas com XXXn devem conter informações que serão incluídas nos modelos em word. 

## 3-	Passo a Passo

Para executar a automação pressione CRTL mais o número da opção que deseja executar:

1.	Consolidar DF Completo

2.	Consolidar DF Modelo 1

3.	Consolidar DF Modelo 2 (Caso haja mais de um modelo)

Por exemplo, para executar a consolidação da DF dos modelos 1 e 2, deve executar o comando “CTRL + 1” dentro do arquivo “ConsolidacaoDF_VF.docm”.
Os resultados da consolidação dos fundos serão gerados dentro da pasta do respectivo Fundo.


