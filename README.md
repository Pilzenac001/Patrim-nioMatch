# Patrimônio Match

O **Patrimônio Match** é uma ferramenta gráfica desenvolvida em Python para automatizar a conciliação patrimonial entre duas bases de inventário distintas: **GEMAT** (Gabinete de Gestão de Materiais) e **SAM** (Sistema de Administração de Materiais).

O objetivo é encontrar, de forma inteligente, o patrimônio correspondente no SAM para cada item listado no GEMAT, gerando um relatório em Excel estilizado, auditável e pronto para o saneamento físico e contábil.

---

## 🚀 Funcionalidades

1. **Leitura Automatizada e Flexível**: Identifica de forma heurística as colunas mesmo que os nomes variem entre as planilhas (ex: "Patrimônio", "Descrição", "Lotação", etc.).
2. **Normalização Avançada**: Limpa termos (remoção de acentos, caracteres especiais, conversão para minúsculas) e converte sinônimos usando um dicionário dinâmico (`sinonimos.json`).
3. **Agrupamento Categorizado**: Reduz o tempo de execução e evita casamentos errôneos agrupando os bens por categorias (ex: armário contra armário).
4. **Cálculo de Confiança por Pesos**: Compara tipo de bem, marca, modelo, medidas, material e similaridade textual total, retornando um índice de 0 a 100%.
5. **Resolução de Conflitos**: Quando mais de um item GEMAT aponta para o mesmo item SAM, a melhor correspondência é mantida e as demais são enviadas para a aba de conferência manual.
6. **Interface Gráfica Premium**: Feita em **CustomTkinter** com suporte a modo escuro/claro e processamento em segundo plano (multi-threading).
7. **Relatório Excel Profissional**: Relatório final com abas separadas por nível de confiança, gráficos e formatação visual moderna (usando `xlsxwriter`).

---

## 🛠️ Estrutura do Projeto

```
PatrimonioMatch/
├── entrada/              # Pasta recomendada para planilhas GEMAT.xlsx e SAM.xlsx
├── saida/                # Pasta de destino dos relatórios gerados
├── logs/                 # Registro em arquivos de logs (opcional)
├── config.json           # Pesos e limites de confiança editáveis
├── sinonimos.json        # Mapeador de termos equivalentes (ex. CPU -> MICROCOMPUTADOR)
├── normalizador.py       # Algoritmos de limpeza e tokenização
├── excel.py              # Leitor de planilhas e mapeador de colunas
├── matcher.py            # Motor principal de similaridade e pesos
├── estatisticas.py       # Funções de cálculo estatístico
├── relatorio.py          # Gerador estilizado de planilhas .xlsx
├── interface.py          # Janela gráfica e navegação
├── main.py               # Ponto de entrada do sistema
├── requirements.txt      # Dependências do Python
└── README.md             # Instruções de instalação e uso
```

---

## 📦 Instalação e Preparação

### 1. Pré-requisitos
Certifique-se de ter o **Python 3.10+** instalado no Windows e a opção `Add Python to PATH` marcada na instalação.

### 2. Criação do Ambiente Virtual
No terminal/cmd, acesse a pasta do projeto e crie o ambiente virtual:
```bash
cd C:\PatrimonioMatch
python -m venv .venv
```

Ative o ambiente virtual:
* **No CMD**:
  ```cmd
  .venv\Scripts\activate
  ```
* **No PowerShell**:
  ```powershell
  .venv\Scripts\Activate.ps1
  ```

### 3. Instalação das Dependências
Com o ambiente ativado, instale as dependências executando:
```bash
pip install -r requirements.txt
```

---

## 💻 Como Utilizar

1. Ative seu ambiente virtual e execute o arquivo `main.py`:
   ```bash
   python main.py
   ```
2. Na janela aberta:
   - Clique em **Procurar...** na seção GEMAT e selecione a planilha de origem.
   - Clique em **Procurar...** na seção SAM e selecione a planilha de referência.
   - Selecione a pasta de saída para salvar o relatório.
   - Ajuste os parâmetros de confiança e pesos se necessário (no painel da direita).
   - Clique em **Iniciar Processamento**.
3. Acompanhe a barra de progresso e as mensagens no painel de logs integrado.
4. Ao finalizar, abra a planilha `Relatorio_PatrimonioMatch.xlsx` gerada na pasta de destino escolhida.

---

## ⚡ Gerar Executável (.exe) para Windows

Caso queira distribuir a aplicação para usuários finais sem a necessidade de instalar Python, você pode compilar em um arquivo executável standalone:

1. Instale o PyInstaller no ambiente virtual (já incluso no `requirements.txt`).
2. No Prompt de Comando/PowerShell, com o ambiente ativo, execute:
   ```bash
   pyinstaller --onefile --windowed --name PatrimonioMatch main.py
   ```
3. O executável standalone será gerado na pasta `dist/` com o nome `PatrimonioMatch.exe`.
4. Copie o executável junto com os arquivos `config.json` e `sinonimos.json` (que devem permanecer na mesma pasta do executável para leitura) e distribua.
