# Higienizador de Modelos Power BI (TMDL + Python)

Este projeto automatiza a remoção de itens não utilizados (medidas, colunas e colunas calculadas) em modelos semânticos do Power BI, utilizando o diagnóstico do **Measure Killer** e manipulando diretamente os arquivos de definição **TMDL**.

## O Problema
Com o tempo, modelos de dados tendem a acumular "dívida técnica": medidas criadas para testes, colunas solicitadas e não usadas, e regras de negócio obsoletas. Isso impacta a performance, a governança e a manutenção do projeto.

## A Solução
O script Python realiza uma "limpeza em ondas", lendo o relatório de uso do Measure Killer e deletando os blocos de código correspondentes nos arquivos `.tmdl` do projeto Power BI (PBIP).

## Regras de Segurança e Governança
Para garantir que a limpeza não corrompa o modelo semântico, o script implementa diversas travas de segurança:
* **Tabelas de Exceção**: Proteção manual para tabelas críticas, como a `d_calendario`.
* **Blindagem de Sistema**: Ignora automaticamente tabelas automáticas de data (`LocalDateTable` e `DateTableTemplate`).
* **Hierarquia de Objetos**: Identifica o início e fim de blocos `column`, `measure` e `hierarchy` para garantir exclusões precisas.
* **Conservadorismo com Colunas**: Remove apenas colunas puramente não utilizadas (`Unused`), preservando aquelas que possam ter dependências de relacionamento.

## Como Utilizar
1. **Diagnóstico**: Execute o Measure Killer e exporte o resultado para um arquivo Excel nomeado `resultado_measure_killer.xlsx`.
2. **Preparação**: Feche o Power BI Desktop e faça um backup da sua pasta de projeto `.PBIP`.
3. **Execução**: Rode o script:
   ```bash
   python limpador.py