# Interpretador PCO - Streamlit

App para interpretar cenários de Controladoria/PCO no padrão RF, RC e RP/NCG.

## Como rodar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como usar

1. Envie uma planilha `.xlsx` ou `.xlsm`.
2. O app tenta localizar automaticamente contas como Faturamento, CPV, Despesas, Empréstimos, NCG, CDG etc.
3. Confira/ajuste a tabela de variações.
4. Escolha o tipo de premissa.
5. Gere respostas prontas para perguntas parecidas com as da prova.

Caso a planilha venha com outro layout, use o modo manual e digite Original/Novo Valor.
