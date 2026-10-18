# Semana do Lar

Agenda semanal da casa em três abas: **Atribuição** (quem deveria fazer cada tarefa), **Execução** (quem fez) e **Consolidado** (como foi o cumprimento da semana).

Página: https://lucasniemeyer13.github.io/semana-da-casa/

Na primeira abertura a página carrega como ponto de partida a semana com as atribuições definidas em 18/out/2026.

## Planilha

A página lê e grava a planilha [Semana da Casa](https://docs.google.com/spreadsheets/d/1v3_Ag52cz6paAL6JA9Tg82lY1sz_yJzb5JKtVDXCItw/edit) (abas Tarefas, Eventos, Histórico e Alterações) por um app da web do Apps Script, código em [`apps-script/Code.gs`](apps-script/Code.gs).

- Cada alteração, e o botão **Salvar**, grava o estado na planilha.
- Ao abrir, ao voltar para a aba e a cada 45 segundos, a página lê a planilha, inclusive o que foi editado direto nela.
- Toda leitura e gravação exige a senha da casa, digitada uma vez em cada navegador. Sem ela o app da web não lê nem grava nada.
