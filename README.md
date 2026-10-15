# Escala da Igreja

Projeto pessoal em Python que gera, de uma vez, a escala de tarefas de todos os domingos de um mês, para os homens e para a igreja em geral, e exporta tudo para uma planilha Excel.

Nasceu de um problema real: montar essa escala à mão, todo mês, era demorado e acabava sobrecarregando sempre as mesmas pessoas, porque cada um tem restrições diferentes sobre o que pode fazer.

> Todos os nomes usados neste repositório são fictícios.

## O que o programa faz

- Gera duas escalas para cada domingo do mês:
  - **Escala dos Homens:** Escola Bíblica Dominical, Oração dos Homens, Louvores, Avisos, Slides, Momento de Oração, Culto e Ceia.
  - **Escala da Igreja:** Pão (café da manhã), Pré-Culto, Almoço e Pós-Culto, organizados por duplas e trios.
- Respeita regras de cada pessoa e de cada tarefa (tarefas fixas, tarefas exclusivas de um grupo, pessoas limitadas a uma tarefa).
- Sorteia com **tentativa e erro (backtracking)**: se a combinação trava no meio do caminho, o programa volta atrás e tenta outra até a escala fechar.
- Dá **folga justa**: toda semana pelo menos um homem descansa, e o rodízio prioriza quem folgou menos.
- Liga as duas escalas: quem não serve na escala dos homens leva o seu grupo para o Pós-Culto, a tarefa mais leve.
- Guarda **histórico entre os meses** em arquivos JSON, para evitar repetir a mesma pessoa na mesma tarefa.
- Dá mais peso aos trios que às duplas no sorteio (configurável).
- Exporta uma planilha Excel com uma aba para cada escala.

## Como rodar

Requisitos: Python 3.8 ou superior.

```bash
pip3 install -r requirements.txt
python3 escala.py
```

Sem argumentos, o programa gera o mês seguinte ao atual. Também é possível escolher o mês:

```bash
python3 escala.py novembro
python3 escala.py novembro 2027
python3 escala.py 11 2027
```

O resultado aparece no terminal e é salvo em `escala_NOMEDOMES_ANO.xlsx`, na mesma pasta.

## Exemplo de saída

```
Semana 1 — 04/10
----------------------------------------
Escala dos Homens:
  Escola Bíblica Dominical     Pedro
  Oração dos Homens            Lucas
  Louvores                     Carlos
  Avisos                       Daniel
  Slides                       Marcos
  Momento de Oração            Gabriel
  Culto                        Pedro
  Ceia                         Rafael
  Folga designada: Eduardo
  Ociosos (vão pro Pós-Culto): André, Eduardo, Felipe
Igreja:
  Pão                          Lucas
  Pré-Culto                    Marcos, Clara e Gabriel
  Almoço                       Lucas e Júlia
  Pós-Culto                    André, Laura e Rafael, Carlos, Sofia e Eduardo, Felipe e Marina
```

## Como configurar

Tudo é configurado nas listas do início do `escala.py`: `HOMENS`, `TAREFAS`, `TAREFAS_FIXAS`, `RESTRICOES_PESSOA`, `RESTRICOES_TAREFA`, `PESSOAS_PAO`, `GRUPOS`, `PESO_TRIO` e `PESO_DUPLA`. Não é preciso mexer na lógica.

## Documentação completa

O passo a passo detalhado de como cada parte funciona, como editar o projeto e como resolver problemas comuns está em [documentacao_escala_igreja_publico.pdf](documentacao_escala_igreja_publico.pdf).

## Tecnologias

- Python 3
- [openpyxl](https://openpyxl.readthedocs.io/) para gerar a planilha Excel
- JSON para guardar o histórico entre os meses

## Arquivos gerados pelo programa

`estado_homens.json`, `estado_igreja.json` e as planilhas `.xlsx` são criados ao rodar o programa e guardam o histórico real de quem já serviu. Por isso eles **não** fazem parte do repositório.
