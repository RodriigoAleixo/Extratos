# 📑 Automação de Geração de Extratos no Sistema NewCon

Este projeto automatiza a geração e o salvamento de extratos de consorciados no sistema **NewCon**, utilizando Python e PyAutoGUI para simular interações humanas com o sistema.

A automação lê uma planilha contendo os dados de **grupo** e **cota**, acessa o site do NewCon, navega até a área de **Extrato do Consorciado** e salva o extrato em formato PDF, nomeando o arquivo com as informações do cliente.

---

## 🚀 Funcionalidades

- Abertura automática do navegador e acesso ao site do NewCon
- Leitura de dados de grupo e cota a partir de uma planilha Excel
- Navegação até o módulo de Extrato do Consorciado
- Geração e salvamento automático do extrato em PDF
- Renomeação dos arquivos com base no grupo e cota

---

## 🧰 Tecnologias Utilizadas

- [Python 3.x](https://www.python.org/)
- [PyAutoGUI](https://pyautogui.readthedocs.io/en/latest/) – Automação de interface
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) – Leitura de arquivos Excel
- [Time](https://docs.python.org/3/library/time.html) – Controle de tempo para sincronização das ações

---

## 📂 Estrutura Esperada da Planilha

A planilha Excel (`extratos.xlsx`) deve conter uma aba com os seguintes campos:

| Grupo | Cota |
|--------|------|
| 123    | 456  |

---

## ▶️ Como Usar

1. Instale as dependências:

```bash
pip install pyautogui openpyxl
```

2. Certifique-se de que o sistema NewCon está acessível e configurado para login automático (se necessário).

3. Ajuste os tempos e posições dos cliques conforme a resolução da tela e comportamento do sistema.

4. Execute o script Python:

```bash
python Extratos.py
```

---

## ⚠️ Observações

- Não mova o mouse durante a execução para evitar interferências.
- A automação depende de reconhecimento visual. Verifique se os elementos visuais (botões, menus) estão no mesmo lugar nas capturas de tela.
- Certifique-se de que o navegador está configurado para imprimir diretamente em PDF.

---

👨‍💻 Autor Rodrigo de Lima Aleixo 
💼 [LinkedIn](https://www.linkedin.com/in/rodrigo-de-lima-aleixo-850b1720b/)
✉️ E-mail: *rodriigoaleixo@gmail.com*

---

Feito para simplificar tarefas repetitivas e ganhar tempo no dia a dia! 🤖✨

