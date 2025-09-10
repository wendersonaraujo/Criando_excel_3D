# 📊 Gerador de Gráfico 3D em Excel com Python (openpyxl)

Este projeto cria uma planilha Excel contendo dados de **lucros e custos por ano** e gera automaticamente um **gráfico de área 3D** utilizando a biblioteca [openpyxl](https://openpyxl.readthedocs.io/).

## 🚀 Funcionalidade
O script:
1. Cria um arquivo Excel (`area.xlsx`);
2. Insere uma tabela de dados com anos, lucros e custos;
3. Gera um gráfico de área 3D comparando **Lucro x Custos**;
4. Salva a planilha com o gráfico embutido.

## 📂 Estrutura dos Dados
Os dados utilizados são porcentagens de lucro e custo por ano:

| Ano  | Lucro (%) | Custos (%) |
|------|-----------|------------|
| 2015 | 30        | 40         |
| 2016 | 25        | 40         |
| 2017 | 30        | 50         |
| 2018 | 10        | 30         |
| 2019 | 5         | 25         |
| 2020 | 10        | 50         |

## 🖼️ Gráfico
O gráfico criado é do tipo **Área 3D (AreaChart3D)**:
- **Título:** Lucros x Custos  
- **Eixo X:** Ano  
- **Eixo Y:** Porcentagem (%)  

O gráfico é inserido na célula `A10` da planilha.

## 🛠️ Requisitos
- Python 3.x  
- Biblioteca [openpyxl](https://pypi.org/project/openpyxl/)

Instalação da biblioteca:
```bash
pip install openpyxl
▶️ Como Executar
Salve o código em um arquivo Python, por exemplo grafico_area.py;

Execute o script:

bash
Copiar código
python grafico_area.py
Será gerado o arquivo area.xlsx na mesma pasta;

Abra o arquivo no Excel para visualizar os dados e o gráfico.

📌 Observação
O gráfico só pode ser visualizado em softwares que suportam gráficos do Excel, como o Microsoft Excel.
No LibreOffice Calc, os dados aparecem, mas o gráfico pode não ser exibido corretamente.

✍️ Autor: Código gerado como exemplo de uso do openpyxl para manipulação de planilhas e gráficos no Excel.

css
Copiar código

Quer que eu também gere uma **explicação resumida** (em poucas linhas) para você usar como legenda em postagem de rede social?







Perguntar ao ChatGPT
