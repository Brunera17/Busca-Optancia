# Busca de Optância por CNPJ

Este projeto é um **script em Python** que consulta automaticamente a **situação de optância de empresas** (Simples Nacional e MEI) a partir de uma lista de **CNPJs em uma planilha Excel**.

O script:
- Lê um arquivo Excel de entrada  
- Identifica automaticamente a coluna que contém o CNPJ  
- Consulta a API pública **ReceitaWS**  
- Trata erros de rede e limites da API com *backoff* progressivo  
- Gera um novo Excel com os resultados  

---

## 📌 Funcionalidades

- 📄 Leitura de CNPJs a partir de arquivo `.xlsx`
- 🔎 Detecção automática da coluna de CNPJ
- 🧹 Padronização e validação de CNPJ
- 🌐 Consulta à API da ReceitaWS
- ⏳ Controle de requisições com atraso automático
- 📊 Geração de planilha de saída com:
  - Nome da empresa
  - CNPJ
  - Optante do Simples Nacional
  - Optante do MEI

---

## 📂 Estrutura de Arquivos
  📁 projeto
  ├── entrada.xlsx # Planilha com os CNPJs
  ├── saida.xlsx # Planilha gerada com os resultados
  └── main.py # Script principal

---

## 📥 Arquivo de Entrada (`entrada.xlsx`)

- Deve conter **uma coluna com CNPJs**
- O nome da coluna pode ser qualquer um que contenha a palavra `cnpj`  
  (ex: `CNPJ`, `cnpj_empresa`, `Cnpj Cliente`)

### Exemplo:

| CNPJ |
|------|
| 12.345.678/0001-90 |
| 98765432000100 |

---

## 📤 Arquivo de Saída (`saida.xlsx`)

Gerado automaticamente com o seguinte formato:

| nome da empresa | cnpj | simples_nacional | mei |
|-----------------|------|------------------|-----|
| Empresa X | 12345678000190 | SIM | NÃO |

---

## 🛠️ Requisitos
---
- Python **3.8 ou superior**
- Bibliotecas Python:
  - `pandas`
  - `requests`
  - `openpyxl`

### Instalação das dependências

```bash
pip install pandas requests openpyxl
```

---
##▶️ Como Executar
---
  Coloque o arquivo entrada.xlsx na mesma pasta do script

  Execute o script:

  python main.py

  Ao final, o arquivo saida.xlsx será criado automaticamente

---
##⚠️ Observações Importantes
---
  A API utilizada (ReceitaWS) possui limite de requisições

  O script aplica atrasos automáticos para evitar bloqueios

  CNPJs inválidos são ignorados

  Em caso de erro de rede ou API, o script tenta novamente

---
##📈 Possíveis Melhorias Futuras
---
  Suporte a API paga (maior volume de consultas)

  Interface gráfica

  Log em arquivo

  Paralelismo controlado

---
##📄 Licença
---
  Projeto de uso livre para fins educacionais e profissionais.
