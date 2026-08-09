# Diário de Bordo — Setup

Dois apps, integrados por planilha: seu app de prospecção continua exatamente como está;
este é um segundo app, com sua própria planilha, publicado como uma API separada.

## Parte 1 — Criar a planilha e publicar a API

1. Acesse [sheets.google.com](https://sheets.google.com) e crie uma planilha nova.
   Renomeie para **"Diário de Bordo"**.
2. No menu, vá em **Extensões > Apps Script**.
3. Apague o conteúdo padrão do editor e cole todo o conteúdo do arquivo `Code.gs`
   (está junto com esta entrega).
4. Salve (ícone de disquete ou Ctrl+S).
5. No topo do editor, no seletor de funções (ao lado do botão "Executar"), escolha
   **setupSheet** e clique em **Executar**.
   - Na primeira vez, o Google vai pedir autorização. Clique em **Revisar permissões**,
     escolha sua conta, clique em **Avançado** > **Acessar Diário de Bordo (não seguro)** e **Permitir**.
   - Isso cria as abas `Biblioteca`, `Log`, `Metricas` e `Config`, já com as 24 tarefas
     padrão do programa.
6. Clique em **Implantar > Nova implantação**.
   - Tipo: clique na engrenagem e escolha **App da Web**.
   - Executar como: **Eu (seu e-mail)**.
   - Quem pode acessar: **Qualquer pessoa**.
   - Clique em **Implantar**, autorize de novo se pedir.
7. Copie a **URL do app da Web** (termina em `/exec`). Você vai usar essa URL no app.

Sempre que editar o `Code.gs`, é preciso ir em **Implantar > Gerenciar implantações**,
clicar no lápis (editar) e escolher **Nova versão** para as mudanças valerem.

## Parte 2 — Publicar o app (PWA)

Igual aos seus outros apps no GitHub Pages:

1. Crie um repositório novo no GitHub (ex: `diario-de-bordo`).
2. Suba os arquivos `index.html`, `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png`.
3. Em **Settings > Pages**, ative o GitHub Pages na branch `main`, pasta raiz.
4. Acesse a URL gerada (ex: `https://seuusuario.github.io/diario-de-bordo/`).
5. Na primeira abertura, o app vai pedir a **URL da API** — cole a URL do passo 7 da
   Parte 1. Fica salva no navegador; só precisa colar uma vez por dispositivo.
   Se precisar trocar depois, use o botão ⚙ no canto inferior direito.
6. No celular, use "Adicionar à tela inicial" no navegador para instalar como app.

## Parte 3 — Integrar com o CRM (opcional, quando quiser)

Como são planilhas separadas, a forma padrão do Google Sheets de puxar dados de uma
planilha para outra sem duplicar nada é a fórmula `IMPORTRANGE`. Ela funciona bem
para ler números prontos (contagem de leads do dia, visitas etc.) sem que eu precise
acessar ou alterar sua planilha de CRM.

Na aba **Metricas** (ou numa aba nova, ex: `Sync`), numa célula vazia:

```
=IMPORTRANGE("URL_DA_SUA_PLANILHA_DE_CRM"; "NomeDaAba!A1:Z1000")
```

- Na primeira vez, o Sheets pede para **permitir acesso** entre as duas planilhas —
  autorize uma vez e fica valendo.
- Depois disso dá para usar `COUNTIFS`/`SUMIFS` em cima do intervalo importado para,
  por exemplo, contar quantos leads Bondmann ou Bransales foram cadastrados hoje e
  jogar esse número automaticamente nos campos "visitas" da semana.

Se quiser, me manda (só o nome das abas e colunas, não precisa compartilhar a
planilha) que eu escrevo a fórmula exata pronta para colar.

## Estrutura da planilha "Diário de Bordo"

| Aba | Colunas | Uso |
|---|---|---|
| `Biblioteca` | id, day, pillar, portfolio, title, desc, ativo | as tarefas que se repetem por dia da semana |
| `Log` | date, taskId, done, timestamp | o que foi marcado como feito, dia a dia |
| `Metricas` | week, visitas, propostas, fechamentos, faturamento, obs | KPIs semanais |
| `Config` | key, value | data de início do programa etc. |
| `Visitas` | id, date, cliente, empresa, portfolio, endereco, objetivo, status, notas | agenda semanal de visitas |
| `Clientes` | id, empresa, cnpj, cnae, segmento, portfolio, contato, telefone, cidade, potencial, notas | base de clientes p/ sugestão de estratégia |
| `Conhecimento` | id, pillar, portfolio, title, content, tags | guias validados (SPIN, Challenger, marca pessoal etc.) |

Tudo isso é uma planilha comum — dá pra abrir, olhar os números e até editar
diretamente lá, além de editar pelo app.

## Atualizando para a versão com Agenda, Clientes e Guia

1. No editor do Apps Script (Extensões > Apps Script), apague todo o conteúdo
   e cole o `Code.gs` novo por cima do antigo.
2. Rode `setupSheet` de novo. É seguro: ele recria `Biblioteca` e `Conhecimento`
   do zero (biblioteca de tarefas e guias padrão), mas **não apaga** nada que
   já exista em `Visitas`, `Clientes` ou `Log`.
3. Vá em **Implantar > Gerenciar implantações**, clique no lápis (editar) na
   implantação existente, em "Versão" escolha **Nova versão** e clique em
   **Implantar**. A URL da API continua a mesma — não precisa colar de novo no app.
4. No repositório do GitHub, suba o `index.html` novo por cima do antigo
   (Add file > Upload files, arrasta e substitui). O GitHub Pages atualiza
   sozinho em cerca de 1 minuto.

## Como importar sua planilha de visitas da semana

No app, aba **Agenda > Importar**: copie da sua planilha de visitas as colunas,
nesta ordem — **Data, Cliente, Empresa, Portfólio, Endereço, Objetivo** — sem
o cabeçalho, cole na caixa de texto e clique em Importar. Cada linha vira uma
visita na Agenda, e as de hoje aparecem automaticamente na aba **Hoje**.

Se preferir manter as duas planilhas conectadas por fórmula em vez de colar
toda semana, use `IMPORTRANGE` (mesmo princípio da seção anterior) numa aba
extra da planilha de visitas, apontando para a aba `Visitas` desta planilha.

## Como funciona a sugestão de estratégia (aba Clientes)

Não usa nenhuma API externa nem chave paga — é um cruzamento de regras dentro
da própria planilha: o `segmento`/`cnae` do cliente é comparado com as tags de
cada guia da aba `Conhecimento`, e os guias mais relevantes (e do portfólio
certo, Bondmann ou Bransales) aparecem como sugestão. Quanto mais completo o
campo *segmento* do cliente (ex: "metalúrgica", "transportadora de cargas"),
melhor a sugestão. Os guias em si ficam na aba `Conhecimento` — edite ou
adicione linhas lá (ou peça pra mim) para deixar as sugestões mais afiadas
com o tempo.
