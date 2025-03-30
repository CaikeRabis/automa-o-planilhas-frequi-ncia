# 📊 Consolidador de Frequência - Google Apps Script

Este projeto automatiza a consolidação de dados de frequência de múltiplas planilhas do Google Sheets, registrando presença em diferentes dias e turnos. A planilha final apresenta:

✅ Nome dos participantes  
✅ Matrícula  
✅ Ligação com a JMU  
✅ Frequência individual por dia/turno  
✅ Percentual de presença  

---

## 🚀 Funcionalidades
- **Automação**: Coleta e unifica dados de múltiplas planilhas.
- **Remoção de duplicatas**: Cada pessoa só é contabilizada uma vez por planilha.
- **Contagem e percentual de frequência**: Converte a contagem de presenças em percentual.
- **Marcação de presença**: Indica os dias e turnos em que a pessoa esteve presente.
- **Ordenação alfabética**: Nomes organizados automaticamente.

---

## 📌 Estrutura da Planilha Consolidada

| Nome         | Matrícula | JMU  | 24/03 (Manhã) | 24/03 (Tarde) | 25/03 (Manhã) | 25/03 (Tarde) | 26/03 (Manhã) | 26/03 (Tarde) | % Presença |
|-------------|-----------|------|--------------|--------------|--------------|--------------|--------------|--------------|------------|
| João Silva  | 123456    | Sim  | X            |              | X            | X            |              | X            | 67%        |
| Maria Souza | 789012    | Não  | X            | X            | X            |              | X            | X            | 83%        |

**Legenda:** ❌ **"X" para presença** e **campo vazio para falta**.

---

## 📂 Como Usar o Script

1. Abra o **Google Planilhas** e vá até **Extensões → Apps Script**.
2. Copie e cole o código do arquivo `consolidarDados.gs`.
3. Modifique as URLs das planilhas no array `urls`.
4. Execute o script para gerar a aba **Consolidado** com os dados formatados.

---

## 📜 Requisitos
- Acesso às planilhas de origem.
- Autorização para executar scripts no Google Apps Script.

---

## 🔗 Tecnologias Utilizadas
- **Google Apps Script** (JavaScript para Google Sheets)
- **Google Sheets API**
