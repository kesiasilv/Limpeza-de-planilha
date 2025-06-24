# 🧼 Limpeza de Planilha Excel com Python

Este projeto contém um script Python que remove **linhas totalmente vazias** de uma planilha Excel, preservando formatação, fórmulas e conteúdos adicionais.

---

## ✅ O que o script faz

- Remove **somente** as linhas **completamente vazias** da planilha.
- **Mantém a formatação original** (cores, bordas, fontes, alinhamento, etc).

---

## 📦 Dependências

Antes de rodar o script, instale os pacotes necessários:
(no terminal adicione esse comando para instalar os pacotes)

```
python3 -m pip install pandas openpyxl
```

## 📚 Explicação das dependências:
pandas: usado em versões alternativas do script (não obrigatoriamente neste).
openpyxl: leitura e manipulação avançada de planilhas .xlsx preservando estilo.

## 🚀 Como rodar o script

- Copie a planilha Excel para a mesma pasta onde está o script Python.
Exemplo de nome: planilha.xlsx
- substitua o nome no código para o nome da sua planilha
```
## alteracao nessa linha do codigo
input_filename = 'planilha.xlsx'
```
- Abra o terminal, navegue até a pasta do projeto e execute:
```
python3 tratar_planilha.py
```
- O script criará ou sobrescreverá o arquivo planilha_tratada.xlsx com as linhas vazias removidas.

## 📝 Observações

- O script funciona apenas com arquivos no formato .xlsx.
- Não suporta arquivos .xlsb, .xlsm, ou .csv diretamente.
- Certifique-se de que a planilha não esteja aberta no Excel ao executar o script, para evitar erro de acesso.

## 👩‍💻 Autores
<div style="text-align: center;"> 
  <a href="https://github.com/kesiasilv">Késia Silva</a>
</div>
<div style="text-align: center;">
  <a href="https://github.com/keitiely">Keitiely Silva</a>
</div>
