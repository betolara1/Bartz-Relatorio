# Relatório - Monitoramento e Análise de XMLs

Sistema de monitoramento e extração de dados de arquivos XML de pedidos, realizando o cruzamento automático com uma tabela de modelos de referência e consolidando as informações em planilhas Excel (`.xlsx`).

---

## 📌 Visão Geral

O projeto foi desenvolvido para automatizar a leitura de arquivos XML gerados em diretórios de rede, extrair as referências de itens (`REFERENCIA`) e compará-las com uma tabela base de modelos (`RELACAO-MODELOS.xlsx`). As quantidades de itens por grupo/ambiente/categoria são contabilizadas e exportadas continuamente para uma planilha consolidada de relatórios.

O repositório conta com:
1. **Aplicação Principal (Java)**: Monitor em tempo real via `WatchService` com persistência contínua no Excel via Apache POI.
2. **Automação de Execução (Batch + VBS)**: Scripts para manter o serviço Java rodando em loop infinito em segundo plano (background).
3. **Script Utilitário / Análise em Lote (Python)**: Processador alternativo em lote usando `openpyxl`, com estilização visual e divisão por categorias/ambientes.

---

## 📂 Estrutura de Arquivos

```
Relatorio_Andre/
├── libs/                     # Dependências JAR (Apache POI, Commons, Log4j, etc.)
├── Main.java                 # Código principal em Java: WatchService e parser DOM XML
├── ExportadorExcel.java      # Módulo Java para leitura e gravação/atualização do Excel
├── RELACAO-MODELOS.xlsx      # Tabela base com os códigos, descrições e categorias
├── executar_loop.bat         # Script Batch que executa o Java e reinicia em caso de queda
├── relatorio.vbs             # Execução silenciosa em background (sem janela CMD)
├── main.py                   # Script alternativo em Python para processamento em lote
└── pom.xml                   # Configuração de referência Maven
```

---

## ⚙️ Fluxo de Funcionamento (Java)

1. **Carregamento da Base de Modelos**:
   - Lê `RELACAO-MODELOS.xlsx` (mapeando a coluna `Item` e o respectivo `Grupo/Categoria`).
2. **Verificação de Histórico**:
   - Lê a planilha de saída consolidada (`relatorio_atualizado.xlsx`) para obter a lista de pedidos já processados, evitando duplicidades.
3. **Carga Inicial de Arquivos**:
   - Varre arquivos `.xml` pré-existentes na pasta configurada (`\\192.168.1.10\DatabaseFolder\01_XML\IMPORTADOS`).
   - Processa os novos pedidos e anexa ao Excel.
4. **Monitoramento em Tempo Real (`WatchService`)**:
   - Monitora eventos de criação (`ENTRY_CREATE`) no diretório de rede.
   - Ao identificar um novo XML, aguarda a finalização da cópia, analisa as tags `<ITEM REFERENCIA="...">`, agrupa por categoria e adiciona os dados à planilha de saída.

---

## 🚀 Como Executar

### 1. Aplicação Java (Padrão de Produção)

#### Pré-requisitos
- **Java JDK/JRE 8+** instalado e configurado no `PATH`.
- Acesso à rede e permissão de leitura/escrita no compartilhamento mapeado (`\\192.168.1.10\DatabaseFolder\...`).

#### Compilação
No diretório do projeto:
```cmd
javac -cp "libs/*;." Main.java ExportadorExcel.java
```

#### Execução Manual (Terminal)
```cmd
java -cp "libs/*;." Main
```

#### Execução com Reinício Automático (Loop)
Execute o arquivo:'
```cmd
executar_loop.bat
```
*(Caso o processo seja interrompido por oscilação de rede ou erro inesperado, o script aguarda 5 segundos e o reinicia automaticamente).*

#### Execução Silenciosa (Segundo Plano)
Dê um duplo clique em:
```cmd
relatorio.vbs
```
*(Inicia o `executar_loop.bat` de forma oculta, sem exibir janela do Prompt de Comando).*

---

### 2. Script Python (Alternativa / Análise em Lote)

#### Pré-requisitos
- **Python 3.8+**
- Biblioteca `openpyxl`:
  ```bash
  pip install openpyxl
  ```

#### Execução
Coloque os arquivos `.xml` desejados no diretório ou aponte a pasta e execute:
```bash
python main.py
```
O script gerará o arquivo consolidado e formatado `RELATORIO_MODELOS.xlsx`.

---

## 🔧 Configurações

Caso seja necessário alterar os caminhos de rede ou de arquivos, edite as constantes no início de [Main.java](file:///c:/Users/Ralf/Desktop/Programa%C3%A7%C3%A3o/Relatorio_Andre/Main.java#L29-L31):

```java
private static String caminho_pasta = "\\\\192.168.1.10\\DatabaseFolder\\01_XML\\IMPORTADOS";
private static String caminho_arquivo = "RELACAO-MODELOS.xlsx";
private static String caminho_saida = "\\\\192.168.1.10\\DatabaseFolder\\relatorio_atualizado.xlsx";
```
*(Lembre-se de recompilar após alterar).*
