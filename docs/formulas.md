# Todas as fórmulas da planilha

Lista gerada a partir da planilha `planilha/excel-formulas-na-pratica.xlsx`. Cada fórmula foi recalculada no Excel 365 em português (26/10/2026) e deu o resultado mostrado. Dados fictícios.

> Excel em português usa **ponto e vírgula** (;) entre argumentos. No Excel em inglês, troque por vírgula e use o nome da função em inglês.

## Aula 1 - Limpeza

### Limpar CPF

| Célula | Fórmula | Resultado |
|---|---|---|
| B2 | `=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(ARRUMAR(A2);".";"");"-";"");" ";"")` | 00000000100 |
| C2 | `=NÚM.CARACT(B2)` | 11 |
| D2 | `=SE(C2=11;"OK";"Conferir")` | OK |
| B3 | `=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(ARRUMAR(A3);".";"");"-";"");" ";"")` | 00000000200 |
| D5 | `=SE(C5=11;"OK";"Conferir")` | Conferir |

### Padronizar nomes

| Célula | Fórmula | Resultado |
|---|---|---|
| G2 | `=MAIÚSCULA(ARRUMAR(F2))` | MARIA EXEMPLO |
| H2 | `=PRI.MAIÚSCULA(ARRUMAR(F2))` | Maria Exemplo |
| G3 | `=MAIÚSCULA(ARRUMAR(F3))` | JOÃO EXEMPLO |
| H4 | `=PRI.MAIÚSCULA(ARRUMAR(F4))` | Ana Modelo |

### Zeros à esquerda

| Célula | Fórmula | Resultado |
|---|---|---|
| K2 | `=TEXTO(J2;"00000000000")` | 00123456789 |

## Aula 2 - Datas

### Cinco dias úteis do curso

| Célula | Fórmula | Resultado |
|---|---|---|
| C4 | `=DIATRABALHO(B4;4;$K$4:$K$5)` | 23/10/2026 |
| D4 | `=DIATRABALHOTOTAL(B4;C4;$K$4:$K$5)` | 5 |
| E4 | `=TEXTO(B4;"dddd")` | segunda-feira |
| C6 | `=DIATRABALHO(B6;4;$K$4:$K$5)` | 05/11/2026 |
| D6 | `=DIATRABALHOTOTAL(B6;C6;$K$4:$K$5)` | 5 |

### Vencimento e situação

| Célula | Fórmula | Resultado |
|---|---|---|
| G4 | `=DATAM(F4;24)` | 06/09/2026 |
| H4 | `=G4-$B$1` | -40 |
| I4 | `=SE(H4<0;"Vencido";SE(H4<=60;"Vermelho";SE(H4<=120;"Amarelo";"Verde")))` | Vencido |
| I5 | `=SE(H5<0;"Vencido";SE(H5<=60;"Vermelho";SE(H5<=120;"Amarelo";"Verde")))` | Vencido |
| I6 | `=SE(H6<0;"Vencido";SE(H6<=60;"Vermelho";SE(H6<=120;"Amarelo";"Verde")))` | Vermelho |

### Lista de feriados

| Célula | Fórmula | Resultado |
|---|---|---|
| K4 | `46328` | 02/11/2026 |

## Aula 3 - Lógica (VT)

### Regra do vale-transporte

| Célula | Fórmula | Resultado |
|---|---|---|
| E4 | `=SE(B4="Não";5*ValorDia;SE(C4="5x2";0;D4*ValorDia))` | R$ 100,00 |
| E5 | `=SE(B5="Não";5*ValorDia;SE(C5="5x2";0;D5*ValorDia))` | R$ 0,00 |
| E6 | `=SE(B6="Não";5*ValorDia;SE(C6="5x2";0;D6*ValorDia))` | R$ 60,00 |
| E7 | `=SE(B7="Não";5*ValorDia;SE(C7="5x2";0;D7*ValorDia))` | R$ 100,00 |
| E8 | `=SE(B8="Não";5*ValorDia;SE(C8="5x2";0;D8*ValorDia))` | R$ 40,00 |
| F4 | `="VT "&MAIÚSCULA(A4)&" · "&TEXTO(E4;"R$ #.##0,00")` | VT FULANO DE TAL · R$ 100,00 |
| G6 | `=SES(B6="Não";5*ValorDia;C6="5x2";0;VERDADEIRO;D6*ValorDia)` | R$ 60,00 |

### Constante nomeada

| Célula | Fórmula | Resultado |
|---|---|---|
| I1 | `20` | R$ 20,00 |

## Aula 4 - Busca

### Puxar nome pelo RE

| Célula | Fórmula | Resultado |
|---|---|---|
| B2 | `=PROCX(B1;tbBase[RE];tbBase[Nome];"Não encontrado")` | Ciclano Souza |
| B3 | `=PROCV(B1;tbBase;2;FALSO)` | Ciclano Souza |
| B4 | `=ÍNDICE(tbBase[Nome];CORRESP(B1;tbBase[RE];0))` | Ciclano Souza |
| B5 | `=SEERRO(PROCV(9999;tbBase;2;FALSO);"RE não encontrado")` | RE não encontrado |
| B6 | `=PROCX(B1;tbBase[RE];tbBase[Área])&" · "&PROCX(B1;tbBase[RE];tbBase[Escala])` | ADMINISTRATIVO · 5x2 |

### Tabela base

| Célula | Fórmula | Resultado |
|---|---|---|
| A9 | `1001` | 1001 |
| B9 | `Fulano de Tal` | Fulano de Tal |

## Aula 5 - Cruzar listas

### Quem entrou e quem saiu

| Célula | Fórmula | Resultado |
|---|---|---|
| B4 | `=SE(CONT.SE($C$4:$C$11;A4)=0;"SAIU";"continua")` | SAIU |
| B5 | `=SE(CONT.SE($C$4:$C$11;A5)=0;"SAIU";"continua")` | continua |
| D4 | `=SE(CONT.SE($A$4:$A$11;C4)=0;"NOVO";"já estava")` | já estava |
| D9 | `=SE(CONT.SE($A$4:$A$11;C9)=0;"NOVO";"já estava")` | NOVO |

### FILTRO e resumo

| Célula | Fórmula | Resultado |
|---|---|---|
| F4 | `=FILTRO(C4:C11;CONT.SE(A4:A11;C4:C11)=0;"nenhum")` | 1009 |
| G4 | `=FILTRO(A4:A11;CONT.SE(C4:C11;A4:A11)=0;"nenhum")` | 1001 |
| J4 | `=CONT.SE(D4:D11;"NOVO")` | 3 |
| J5 | `=CONT.SE(B4:B11;"SAIU")` | 3 |
| J6 | `=CONT.SE(B4:B11;"continua")` | 5 |

## Aula 6 - Contar e somar

### Tabela de movimentação

| Célula | Fórmula | Resultado |
|---|---|---|
| A4 | `Fulano de Tal` | Fulano de Tal |

### Contar e somar com condição

| Célula | Fórmula | Resultado |
|---|---|---|
| G4 | `=CONT.SE(tbMov[Status];"ADMITIDO")` | 8 |
| G5 | `=CONT.SE(tbMov[Status];"DEMITIDO")` | 4 |
| G6 | `=CONT.SES(tbMov[Status];"ADMITIDO";tbMov[Área];"OPERACIONAL")` | 6 |
| G7 | `=SOMASE(tbMov[Status];"DEMITIDO";tbMov[Valor])` | R$ 1.633,20 |
| G8 | `=SOMASES(tbMov[Valor];tbMov[Status];"DEMITIDO";tbMov[Área];"LOGÍSTICA")` | R$ 1.543,30 |
| G9 | `=MÉDIASE(tbMov[Status];"ADMITIDO";tbMov[Valor])` | R$ 272,48 |
| G10 | `=SE(G4+G5=CONT.VALORES(tbMov[Nome]);"OK";"Conferir")` | OK |

## Aula 7 - Desconto

### Competência, total e texto

| Célula | Fórmula | Resultado |
|---|---|---|
| E4 | `=TEXTO(B4;"mm/aaaa")` | 10/2026 |
| F4 | `=A4&" parcela: "&TEXTO(C4;"R$ #.##0,00")&" ("&E4&") - valor referente a "&D4` | 1ª parcela: R$ 50,00 (10/2026) - valor referente a material de exemplo |
| C6 | `=SOMA(C4:C5)` | R$ 75,50 |
| F6 | `="Total a descontar: "&TEXTO(C6;"R$ #.##0,00")&" em "&CONT.VALORES(C4:C5)&" parcelas"` | Total a descontar: R$ 75,50 em 2 parcelas |

## Aula 8 - Máscara LGPD

### Mascarar CPF e e-mail

| Célula | Fórmula | Resultado |
|---|---|---|
| B4 | `="***"&EXT.TEXTO(A4;4;9)&"**"` | ***.000.001-** |
| D4 | `=ESQUERDA(C4;1)&"***"&EXT.TEXTO(C4;LOCALIZAR("@";C4);99)` | m***@exemplo.com |
| B6 | `="***"&EXT.TEXTO(A6;4;9)&"**"` | ***.000.003-** |
| D6 | `=ESQUERDA(C6;1)&"***"&EXT.TEXTO(C6;LOCALIZAR("@";C6);99)` | a***@exemplo.com |
