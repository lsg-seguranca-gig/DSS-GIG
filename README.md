# DSS GIG — Diálogo Semanal de Segurança

Site estático + Google Apps Script + Google Sheets.

## Estrutura
- `index.html` — Página do **colaborador** (assistir vídeos, assinar e registrar)
- `gestor.html` — Página do **gestor** (filtros e geração de **PDF** com logo, semana, tabela e rodapé)
- `styles.css` — Estilos visuais
- `apps-script/Code.gs` — Backend (API) em Google Apps Script que lê/escreve na planilha

## Pré-requisitos
- Uma planilha com ID já fornecida por você contendo as abas:
  - `Funcionarios (Matricula, Nome, Setor, Ativo)`
  - `Treinamentos (SemanaISO, Titulo, URL, Ativo)`
  - `Registros (Timestamp, Matricula, Nome, Setor, SemanaISO, TituloVideo, URLVideo, AssinaturaPNG, DeviceInfo)`

## Passo a passo — Backend (Apps Script)
1. Abra a planilha → **Extensões → Apps Script**.
2. Crie um projeto e substitua o conteúdo por `apps-script/Code.gs`.
3. **Implantar** → **Implantar como aplicativo da Web**:
   - Executar como: **sua conta**
   - Quem tem acesso: **Qualquer pessoa com o link**
4. Copie a **URL do aplicativo da Web**.

## Passo a passo — Front-end
1. Abra `index.html` e `gestor.html` e substitua `PASTE_APPS_SCRIPT_WEB_APP_URL_AQUI` pela URL do seu Apps Script.
2. Publique os três arquivos (`index.html`, `gestor.html`, `styles.css`) em um host estático (ex.: **Netlify**).

## Uso
- **Colaborador**: entra em `index.html`, informa a **matrícula**, escolhe um dos **3 vídeos mais recentes ativos**, assiste até o fim, **assina** e **registra**.
- **Gestor**: acessa `gestor.html`, filtra por **Matrícula**, **Nome** ou **Semana (YYYY-Www)** e **gera o PDF** com **logo**, **título**, **semana**, **tabela** e **rodapé** (3 linhas).

## Observações
- O **Timestamp** do registro é gerado no **servidor** (Apps Script) para maior integridade.
- A **assinatura** é salva como imagem **PNG** (data URL) na planilha.
- Os **vídeos** devem ser URLs do YouTube (não listados), um por linha na aba `Treinamentos`.
- Semana no formato `YYYY-Www` (ex.: `2026-W03`).

## Melhorias futuras (opcionais)
- PIN/Senha para o **gestor**.
- Exportação **CSV**/Excel no painel do gestor.
- Suporte a **youtube-nocookie.com** (privacidade reforçada).
