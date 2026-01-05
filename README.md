📄 Extração de Dados de CTe (DACTE)

Projeto em Python para leitura de arquivos PDF de CTe (DACTE) e extração automática de informações operacionais, consolidando os dados em uma planilha Excel para análise e controle logístico.

🎯 Objetivo

Automatizar a leitura de CTes em PDF

Extrair dados relevantes do documento

Reduzir trabalho manual e erros de digitação

Apoiar análises operacionais e administrativas

📌 Informações Extraídas

Data e hora de emissão

Início e término da prestação

Quantidade (QTD / KG)

Valor total da mercadoria

Valor do frete

Número do documento

Motorista

Placas dos veículos

Emissor

Registro de erro para PDFs inválidos

🛠️ Tecnologias

Python

pdfplumber

Pandas

Regex (re)

⚙️ Funcionamento

Leitura automática de todos os PDFs em uma pasta

Extração do texto de cada página

Identificação dos campos via expressões regulares

Tratamento de exceções (PDF corrompido ou inválido)

Consolidação dos dados em Excel

📤 Saída

Arquivo Excel com todos os CTes processados
