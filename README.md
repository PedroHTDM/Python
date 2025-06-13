# 🏦 ETL de Fundos Oracle

Este projeto automatiza a importação de dados de diversas planilhas Excel para uma tabela Oracle. Ele padroniza colunas, trata tipos de dados, exclui dados antigos e insere dados novos. Também envia e-mails em caso de erro durante o processamento.

---

## ✨ Funcionalidades

* Leitura de arquivos Excel com layouts diferentes
* Mapeamento dinâmico de colunas
* Conversão de tipos: string, número e data
* Limpeza e substituição de valores nulos
* Exclusão de dados antigos antes da inserção
* Inserção em lote no Oracle Database
* Envio automático de e-mail em caso de erro

---

## ⚙️ Requisitos

* Python 3.x
* Oracle Instant Client
* Bibliotecas Python:

  * `pandas`
  * `oracledb`
  * `openpyxl`

Instale com:

```bash
pip install pandas oracledb openpyxl
```

---

## 🔄 Execução Manual

### 1. Configure o script `importar_fundos.py`

* Atualize caminhos de planilhas, aba e mapeamentos
* Informe credenciais do Oracle e do e-mail

### 2. Crie o arquivo `.bat`

Crie um arquivo chamado `rodar_importacao.bat` com o seguinte conteúdo:

```bat
@echo off
REM Caminho real do Python e do script corrigido
"C:\Caminho\Para\Python\python.exe" "C:\Caminho\Para\importar_fundos.py"
```

### 🔍 O que cada linha faz:

* `@echo off`: oculta os comandos executados no terminal
* `REM ...`: comentário explicativo dentro do `.bat`
* Primeira linha executável:

  * O primeiro caminho é para o executável do Python (ex: `C:\Python311\python.exe`)
  * O segundo é para o script `.py` (ex: `C:\projetos\etl\importar_fundos.py`)

> Esse `.bat` é direto e útil para ser usado em servidores ou agendado com tarefas automatizadas.

---

## ⏰ Agendamento com o Agendador de Tarefas (Windows)

1. Abra o **Agendador de Tarefas** (`taskschd.msc`)
2. Clique em "Criar Tarefa"
3. Na aba **Geral**:

   * Marque: "Executar com privilégios mais altos"
4. Aba **Disparadores (Triggers)**:

   * Novo > selecione o horário desejado (ex: diariamente às 02:00)
5. Aba **Ações**:

   * Ação: Iniciar um programa
   * Programa: `C:\caminho\para\rodar_importacao.bat`
6. Aba **Condições** e **Configurações**:

   * Desmarque restrições (como "Somente se com energia")
7. Clique em OK e teste a tarefa

---

## 🌐 Estrutura Recomendada do Projeto

```
projeto-etl-oracle/
├── importar_fundos.py
├── rodar_importacao.bat
├── logs/
│   └── execucao_YYYY-MM-DD.log
└── README.md
```

---

## 📄 Licença

Este projeto está licenciado sob a Licença MIT.
