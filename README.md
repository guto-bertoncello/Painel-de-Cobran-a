# Painel de Cobrança

Aplicação web local criada a partir dos requisitos de `Requisitos.txt`, com layout inspirado na referência enviada e foco em:

- importar base principal e base de novação;
- salvar as importações localmente no navegador para leituras futuras;
- apagar apenas a cópia salva pelo site, sem apagar o arquivo original do computador;
- filtrar, deduplicar CPF, exportar `.xlsx`, gerar gráfico de ações e relatório por agente.

## Como abrir

Opção mais simples no Windows:

1. Execute [iniciar-site.bat](/C:/Users/Gustavo/Downloads/Pessoal/Codex/site-cobranca/iniciar-site.bat).
2. O script sobe um servidor local em `http://localhost:8080`.
3. Feche a janela do servidor quando quiser encerrar o site.

Opção manual:

```powershell
py -3 -m http.server 8080 --directory "C:\Users\Gustavo\Downloads\Pessoal\Codex\site-cobranca"
```

Depois abra [http://localhost:8080](http://localhost:8080).

## O que o site faz

- Tela 1: importação, persistência local e exportação da base principal.
- Tela 2: filtros, lista de ações com defasagem mínima/máxima e resumo salvo.
- Tela 3: gráfico de pizza por ação com atualização automática conforme a seleção.
- Tela 4: importação dedicada de novação e gráfico em pizza com valores “em atraso” e “a receber” do mês atual.
- Tela 5: relatório por agente com “Acionamentos” e “Acordos”.

## Arquivos principais

- [index.html](/C:/Users/Gustavo/Downloads/Pessoal/Codex/site-cobranca/index.html)
- [assets/styles.css](/C:/Users/Gustavo/Downloads/Pessoal/Codex/site-cobranca/assets/styles.css)
- [assets/app.js](/C:/Users/Gustavo/Downloads/Pessoal/Codex/site-cobranca/assets/app.js)
- [assets/logic.js](/C:/Users/Gustavo/Downloads/Pessoal/Codex/site-cobranca/assets/logic.js)

## Observações

- A exportação remove as colunas D, I, K, M, N, O, P, Q e R, conforme os requisitos.
- A divisão por operadores gera exatamente a quantidade de abas definida no filtro.
- Os arquivos de teste originais em `C:\Users\Gustavo\Downloads\Pessoal\Codex` não são apagados pelo site.

## Testes da lógica

```powershell
"C:\Users\Gustavo\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --test "C:\Users\Gustavo\Downloads\Pessoal\Codex\site-cobranca\tests\logic.test.mjs"
```
