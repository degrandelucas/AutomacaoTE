# Automatização de Tarefas na Plataforma de Tracking

Este projeto consiste na automatização de tarefas na Plataforma de Tracking da empresa, utilizando Python para criar um script que realiza duas tarefas principais.

## Funcionalidades Principais

1. **Finalização de Notas Fiscais Entregues:** O script importa uma planilha com informações sobre notas fiscais, verifica se estão marcadas como "MERCADORIA ENTREGUE" e, em caso afirmativo, insere a data de entrega na plataforma de tracking.

2. **Adição de Observações em Notas Fiscais Não Entregues:** Para as notas fiscais que não foram entregues, o script adiciona as observações na plataforma de tracking com base nas informações presentes na planilha importada.

## Passo a Passo do Projeto

### Passo 1: Entrar no Sistema da Empresa

O primeiro passo é automatizar o processo de login e acesso à plataforma de tracking.

### Passo 2: Importar a Base com as Informações do TE

Após entrar no sistema, o script importa uma planilha com as informações das notas fiscais e realiza tratamentos necessários nos dados.

### Passo 3: Finalizar as Notas Fiscais Entregues ou Inserir Observações

Para cada nota fiscal na planilha, o script verifica se foi entregue. Se sim, insere a data de entrega na plataforma. Caso contrário, adiciona as observações correspondentes.

## Tecnologias Utilizadas

- **Python:** Utilizado para desenvolver o script de automação das tarefas.
- **PyAutoGUI:** Biblioteca utilizada para realizar automação de tarefas por meio de interações com a interface gráfica.
- **Pandas:** Biblioteca utilizada para manipulação e análise de dados.
- **Plataforma de Tracking:** Sistema utilizado para monitorar o status das notas fiscais.

## Autor

- **Nome:** Lucas Degrande
- **Contato:** lucasdegrande15@gmail.com
