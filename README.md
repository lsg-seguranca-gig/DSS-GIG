# DSS GIG — Diálogo Semanal de Segurança (CommonJS)

Projeto completo (site estático) com **proxy serverless (CommonJS)** na Vercel para evitar CORS.

## Conteúdo
- `index.html` — Colaborador (matrícula → vídeo → assinatura → registro)
- `gestor.html` — Gestor (filtros + geração de PDF com logo/tabela/rodapé)
- `styles.css` — Estilos
- `apps-script/Code.gs` — Backend (Google Apps Script) que integra com a planilha
- `api/gas.js` — **Proxy CommonJS** (sem CORS) que chama o Apps Script

## O que você precisa editar
- Nada no front — o `API_BASE` já aponta para `/api/gas`.
- O `api/gas.js` **já** está com a sua URL do Apps Script `/exec`.

## Rodar em desenvolvimento (recomendado)
```bash
npm i -g vercel
vercel login
vercel dev
# abra http://localhost:3000
```
Teste rápido:
- `http://localhost:3000/api/gas?action=treinamentos` → deve retornar JSON.

## Deploy na Vercel
Configurações do projeto:
- Framework Preset: **Other**
- Root Directory: `./`
- Build Command: **(vazio)**
- Output Directory: `.`

A Vercel publicará os arquivos estáticos e a função `/api/gas` automaticamente.

## Observações
- A assinatura é salva como PNG (data URL) na planilha.
- Nome das abas: `Funcionarios`, `Treinamentos`, `Registros`.
- Em caso de erro, confira os logs do `vercel dev` e acesse diretamente o GAS:
  `https://script.google.com/macros/s/SEU_EXEC/exec?action=treinamentos`.
