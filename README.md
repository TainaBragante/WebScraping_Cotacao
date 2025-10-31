# WebScraping de Cotação de Criptomoedas no Excel

O **Robô de Cotação** é uma automação em **Python + Selenium** que visita páginas de histórico de preços de criptomoedas, captura os **valores de fechamento** e atualiza planilhas de **Excel** individuais para cada ativo, com foco de manter um **histórico confiável**, sem duplicar datas.


No final, você terá:
* 📈 **Planilhas por moeda** com data e preço de fechamento formatados.
* 🧹 **Sem duplicatas:** o robô confere as datas já existentes antes de escrever.
* ⚡ **Execução rápida e estável:** rolagem/teclas de escape para contornar pop-ups e carregamentos.

---

## Funcionalidades

* 🔎 Busca o **fechamento do(s) último(s) dia(s)** indicado(s) em `dias_para_buscar` (ex.: `[1]`, `[1,2,3]`).
* 🗂️ Lê o arquivo **MOEDAS.xlsx** com a lista dos ativos e suas respectivas **URLs** de histórico.
* 📚 Atualiza o arquivo `{NomeDaMoeda}.xlsx` com **Data** e **Valor de Fechamento**.
* 🧾 Aplica estilos: data (`DD/MM/YYYY`) e número (`0.0000`).

---


## Estrutura de pastas sugerida

```
C:\Users\Admin\Cotacao\
├─ MOEDAS\
│  └─ MOEDAS.xlsx              # Lista mestre: nome da moeda (A) e URL da moeda (B)
└─ COTACAO_DOLAR_COINMARKET\
   ├─ BTC.xlsx                  # Planilha por moeda (aba Sheet1)
   ├─ ETH.xlsx
   └─ ...
```

Cada `{Moeda}.xlsx` deve conter a aba **Sheet1** com duas colunas:

1. **A**: Data (a automação escreve como `datetime`, formato `DD/MM/YYYY`)
2. **B**: Fechamento (número, formato `0.0000`)


---

## Como executar

**Clone/baixe** o projeto e crie um ambiente:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   pip install selenium webdriver-manager openpyxl
   python robo_cotacao.py
   ```

* O navegador abrirá em modo automatizado, visitará cada URL da lista e atualizará as planilhas.
* Ao finalizar, o Chrome fecha e a pasta temporária é removida.

