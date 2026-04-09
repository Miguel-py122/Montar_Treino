# Planner de Treinos

Aplicacao web front-end para montar fichas de treino com preenchimento rapido, busca de exercicios por agrupamento muscular, salvamento automatico no navegador e exportacao profissional.

## Recursos

- Tabela interativa com edicao direta
- Treinos separados em abas
- Biblioteca ampliada de exercicios por agrupamento muscular
- Busca pesquisavel com sugestoes automaticas por agrupamento
- Dropdown visual customizado para agrupamento muscular
- Presets de divisao do dia e mapa corporal interativo
- Mapa corporal em SVG com alternancia frente e costas
- Presets com exercicios-base sugeridos automaticamente
- Duplicacao rapida de linhas de exercicio
- Salvamento automatico com LocalStorage
- Exportacao em JSON, CSV e XLSX
- Exportacao XLSX com uma aba por treino e cabecalho formatado
- Importacao de treinos por JSON e CSV
- Layout responsivo

## Como usar

1. Abra o arquivo `index.html` no navegador.
2. Use os presets de divisao do dia ou o mapa corporal para acelerar a montagem.
3. Escolha o agrupamento muscular no dropdown visual e pesquise os exercicios da linha.
4. Use o botao "Duplicar" para repetir uma linha preenchida em um clique.
5. Preencha os campos da linha como em uma planilha.
6. Use "Importar treino" para carregar JSON ou CSV quando precisar.
7. Use "Baixar treino" para exportar em JSON, CSV ou XLSX.

## Estrutura

- `index.html`: estrutura da aplicacao
- `styles.css`: interface e responsividade
- `script.js`: logica, biblioteca de exercicios, armazenamento local, importacao e exportacao
- `exemplos/`: arquivos-modelo para importacao