# Validador de Registros com Regras de Negócio (Excel + VBA)

Cole sua tabela, cadastre suas regras e clique em **Conferir**. A planilha aponta
todos os registros que violam alguma regra e gera uma aba com o nome da regra e a
mensagem de erro de cada um.


## Por que criei
Muitas rotinas de auditoria exigem aplicar vários filtros na mesma tabela, repetidas
vezes, com regras que mudam com frequência. Além disso, nem sempre é permitido enviar
dados sensíveis para ferramentas de IA ou serviços em nuvem. Esta planilha roda 100%
local, no seu Excel.

## Quando usar
- Tabelas com dados sensíveis que não podem ser enviadas a terceiros
- Validações com muitas regras de negócio que mudam com frequência
- Rotinas em Excel que exigem vários filtros sobre a mesma tabela

## Download
Baixe o arquivo mais recente em [**Releases**](../../releases/latest).

> **Macros bloqueadas?** Clique com o botão direito no arquivo → Propriedades →
> marque **Desbloquear** → OK. Depois abra e habilite o conteúdo.

## Como usar
1. Trate os dados (formate datas e valores, corrija espaços, remova mesclagens)
2. Cole a tabela na aba **Tabela**
3. Cadastre as regras na aba **Lista de Regras**
   (lembre-se: você cadastra o que o sistema deve entender como **ERRADO**)
4. Clique em **Conferir**
5. Se houver erros, veja a nova aba criada

## Tipos de regra
| Operador | Aponta erro quando... |
|---|---|
| `=` | o valor é igual ao parâmetro |
| `<>` | o valor é diferente do parâmetro |
| `>` / `<` | o valor é maior / menor (número ou data) |
| `Contém` / `Não Contém` | o texto contém / não contém algum termo |
| `Combinação` | a linha contém todos os termos informados |
| `Encontrar` | o registro esperado não existe na tabela |
| `Filtro` | agrupa registros para validar soma, contagem ou média |

## Para aprender
A planilha acompanha dados fictícios e regras de exemplo. Ative **uma regra por vez**
e compare a aba de erros com a tabela para entender o comportamento de cada uma.

## Documentação
- [Manual de instruções (PDF)](docs/Manual_Instrucoes_Validar_Registro.pdf)
- [Escopo do sistema](docs/Escopo_do_Sistema.md)
- [Código VBA](src/)

## Feedback
Sugestões e problemas são bem-vindos em [Issues](../../issues) ou no meu
[LinkedIn](https://www.linkedin.com/in/caio-henrique-625682a2/).

## Licença
MIT
