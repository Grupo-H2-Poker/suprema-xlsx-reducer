**Redutor da planilha Suprema 113**

Gera uma versão enxuta do relatório semanal ou diário da Suprema (liga 113 / Grand Union), contendo só o que o importador `process_suprema_113.py` usa. Um arquivo de 250 MB cai para cerca de 45 MB em uns 2 minutos.

Este script recria o redutor que já era usado antes, mas cujo código-fonte se perdeu. A saída foi comparada célula a célula com uma planilha reduzida pelo script antigo (semana 21 a 27/09/2026) e é idêntica.

---

**O que ele faz**

- Mantém só 4 abas: `Grand Union Data`, `Grand Union Member Resume`, `Grand Union Game Statistics` e `Grand Union Game Detail`. As outras 12 são descartadas.
- Na aba `Grand Union Game Detail`, mantém só os jogadores dos **clubes H2**. Todos os blocos de jogo continuam na planilha, inclusive os que não têm nenhum jogador H2.
- As outras 3 abas são copiadas inteiras, nas mesmas linhas e colunas do original.
- Remove a formatação, as células vazias e as linhas em branco.

Um clube é considerado H2 quando o nome começa com `H2poker` ou contém `H2 `. É o mesmo critério do PHP `clubesH2()`. Entram, por exemplo, `H2poker 1`, `H2poker | ITM` e `H2 Club Curitiba`. Ficam de fora os parceiros `... ♤ H20`. Para mudar o critério, edite `PATTERN_CLUBE_H2` em `reduz_planilha.py`.

---

**Instalação (uma vez só)**

No Ubuntu/Debian, o `pip` e o `venv` não vêm instalados. É preciso instalar com sudo:

```bash
sudo apt install python3-pip python3.13-venv
```

Depois crie um ambiente virtual dentro da pasta `reduz_planilha` e instale a dependência:

```bash
cd reduz_planilha
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Não use `sudo pip install` nem `pip install` fora do `.venv`. O Ubuntu bloqueia essa instalação com o erro `externally-managed-environment`, e ela pode quebrar pacotes do sistema.

---

**Uso**

Em um terminal novo, ative o ambiente virtual antes de rodar:

```bash
cd reduz_planilha
source .venv/bin/activate
```

Coloque as planilhas em `pendentes/` e rode o script sem argumentos:

```bash
python reduz_planilha.py
```

Para cada planilha em `pendentes/`:

- a versão reduzida é gravada em `processados/` **com o mesmo nome** e já pode ser colocada em `uploads/` para o importador;
- o original é movido para `processados/originais/`;
- se der erro, o arquivo fica em `pendentes/` e o script segue para a próxima planilha.

Só planilhas da liga **113** são reduzidas. As planilhas 106 e 107 são recusadas e ficam em `pendentes/`, porque o importador delas usa as abas `Union *`, que o redutor descartaria.

Também é possível reduzir um arquivo específico, fora das pastas:

```bash
python reduz_planilha.py ENTRADA.xlsx
# gera ENTRADA_reduzido.xlsx na mesma pasta

python reduz_planilha.py ENTRADA.xlsx -o SAIDA.xlsx
```

O script nunca sobrescreve a planilha original.

---

**Exemplo**

A pasta `exemplo/` tem uma planilha pequena para testar o script sem precisar de um relatório real:

- `113-exemplo.xlsx`: entrada no mesmo formato da Suprema, com 5 abas, 3 jogos e poucos jogadores por jogo.
- `113-exemplo_reduzido.xlsx`: saída gerada pelo script a partir da entrada.

Para gerar a saída de novo:

```bash
python reduz_planilha.py exemplo/113-exemplo.xlsx
```

O que dá para conferir na saída:

- a aba `Union Data` foi descartada;
- na `Game Detail`, ficaram só os jogadores dos clubes `H2poker ...`, e os parceiros `... ♤ H20` saíram;
- o jogo `3-MAX TURBO NLH II(8)` continua na planilha, mesmo sem nenhum jogador H2.

Os dados vêm da semana de 21 a 27/09/2026, mas os nomes e IDs de jogadores e agentes foram trocados por valores fictícios (`Jogador 1`, `900001`, `Agente 1`, `800001` etc.).

---

**Detalhes técnicos**

- A leitura é feita direto do XML do `.xlsx`, em streaming, com a biblioteca padrão do Python. O openpyxl seria lento demais para a aba `Game Detail`, que tem mais de 1 GB de XML descompactado.
- A escrita usa o XlsxWriter no modo `constant_memory`, o mesmo do script antigo. Com isso, o uso de memória fica baixo.
- Diferença em relação ao script antigo: textos que começam com `=` (por exemplo, o nome de jogador `=$hum=`) eram gravados como fórmula e viravam `0`. Agora são gravados como texto.
