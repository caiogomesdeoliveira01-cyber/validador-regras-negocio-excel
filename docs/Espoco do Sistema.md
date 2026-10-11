# Botão Conferir:
1º Percorre linha por linha da Tabela, e em cada uma vê se existe uma regra de negócio para validar
2º Caso encontre uma regra, verifica se registro está correto ou errado
3º Se estiver errado, anota todos os dados do registro junto com nome da regra e mensagem de erro
4º Cola em uma nova planilha todos os registros que estão errados conforme as regras de negócio

# Botão Limpar Registros:
- Limpa todos os registros da tabela. Mantendo somente o cabeçalho.

# Botão Excluir Tabela:
- Limpa todos registros e colunas da tabela, deixando-a pronta para colar um novo tipo de tabela.
- Recomenda-se que após colar nova tabela verifique se o tipo de dados de cada coluna está correta.
- Recomenda-se que verifique se continua formatado como tabela e nomeado como tbl_tabela

# Planilha Lista de Regras:
- Local onde é cadastrado todas as regras de negócio
- Só executa regras onde Ativo for igual a Sim
- Caso alguma regra dê erro durante a execução. Na coluna Ativo fica igual a Erro VBA.
- Parte 1 serve para saber se aquela regra se aplica ao registro. Se positivo segue para Parte 2.1, se negativo pula para a próxima regra.
- Parte 2.1 serve para olhar somente o registro atual para ver se está errado. Se estiver com o Filtro ele pula para a parte 2.2. E caso encontre algum registro errado, segue para a Parte 3, e caso negativo pula para a próxima regra.
- Parte 2.2 serve para filtrar todos os registros seguindo as regras da Parte 2.1, depois faz uma conta e por fim compara o resultado com o Valor_Condição2 para saber se todos estão corretos ou errados.
- Parte 3: serve para explicar o que a regra de negócio aplicou e porque o registro está errado (caso encontre).
- Na coluna Valor_Condição1 você pode colocar mais de um valor como parâmetro usando o sinal ; como separador. Também recomenda-se usar Alt + Enter para pular linha dentro da célula e assim ficar mais claro quais são os valores condicionais.


# Colunas da Lista de Regras:
- **Nome da Regra:** nome dado a regra de negócio. Recomenda-se não repetir o nome para evitar interpretação errada de qual regra foi aplicada.
- **Ativo:** onde define se a regra está em funcionamento ou não (sistema só considera olhar regras onde estiver escrito Sim).
- **Campo Alvo:** nome da coluna da tabela que o sistema deve olhar
- **Valor Alvo:** define qual nome o sistema deve encontrar na coluna do Campo Alvo. Caso encontre, sistema entende que para aquela linha deve aplicar a regra para saber se está correto ou errado.
- **Campo_Condição1:** nome da coluna da tabela que o sistema deve olhar
- **Operador_Condição1:** define o tipo de regra que deve analisar
- **Valor_Condição1:** define o valor como parâmetro para saber se está errado. Combinado com o Operador_Condição1, caso encontre um registro na mesma condição, o sistema considera que o registro está errado. A exceção é somente para Operador_Condição1 igual a Filtro, que segue adiante validando de outra forma.
- **Tipo Agrupamento:** operadores especiais que executa somente uma vez separando todos os registros que fazem parte do operador Filtro. Feito isso define o tipo de cálculo deve fazer.
- **Campo_Condição2:** nome da coluna da tabela que o sistema deve olha
- **Operador_Condição2:** define o tipo de comparação que deve analisar
- **Valor_Condição2:** define o valor como parâmetro para saber se está errado. Combinado com o Operador_Condição2 e Tipo Agrupamento, caso o resultado do cálculo esteja na mesma condição, o sistema considera que todos aqueles registros estão errados. Assim basta apenas olhar qual deles realmente está errado.
- **Já foi Agrupado?:**  coluna auxiliar do sistema para validar se alguma regra da Parte 2.2 foi executada. Em caso positivo, desconsidera pois a verificação já foi feita.
- **Mensagem do Erro:** escreve uma mensagem caso uma condicional da Parte 2.1 ou 2.2 seja true. Aqui espera-se que a mensagem colocada fique claro qual o erro que foi encontrado.

# Operador_Condição1:
- **Sinal =** entende que o registro está errado se o valor da coluna Campo_Condição1 for IGUAL ao Valor_Condição1
- **Sinal <>** entende que o registro está errado se o valor da coluna Campo_Condição1 for DIFERENTE ao Valor_Condição1
- **Sinal >** entende que o registro está errado se o valor da coluna Campo_Condição1 for MAIOR ao Valor_Condição1. Caso a coluna não esteja como tipo numérico ou data, sistema vai considerar erro na regra.
- **Sinal <** entende que o registro está errado se o valor da coluna Campo_Condição1 for MENOR ao Valor_Condição1. Caso a coluna não esteja como tipo numérico ou data, sistema vai considerar erro na regra.
- **Contém** entende que o registro está errado se o valor da coluna Campo_Condição1 CONTÉM algum trecho do Valor_Condição1
- **Não Contém** entende que o registro está errado se o valor da coluna Campo_Condição1 NÃO CONTÉM algum trecho do Valor_Condição1
- **Encontrar** sistema procura algum tipo de registro. Caso não encontre significa falta esse registro na tabela, que pode ser porque não foi lançado ou que foi lançado com outro nome.
- **Combinação** sistema percorre todas as colunas do registro e entende que está errado se encontrar todos os valores descritos no Valor_Condição1. Isso é ideal quando precisa validar mais de duas colunas.


# Treinamento do Sistema:
- A planilha já vem uma tabela de registros e uma lista de regras para entender como funciona na prática.
- Os dados são fictícios e usados apenas para exemplificar o funcionamento do sistema.
- Recomenda-se deixar uma regra ativa por vez para saber como funciona na prática
- Há um manual de instruções como apoio.




								
