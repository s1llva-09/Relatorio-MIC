# Relatorio MIC

> **Movimentos Internos Custeados** para o Protheus.

O MIC consulta os movimentos de estoque da tabela `SD3`, relaciona os dados do
produto, solicitacao, funcionario, usuario e tipo de movimento, e gera uma
planilha `.xlsx` pronta para analise.

## O que o relatorio entrega

- Movimentos internos com custo, quantidade e custo unitario.
- Descricao do produto e do tipo de movimento.
- Centro de custo, conta contabil, lote, documento e ordens relacionadas.
- Solicitante, funcionario responsavel e atendente, quando encontrados.
- Exclusao de movimentos estornados, de producao e das transferencias `DE4` e
  `RE4`.

## Formas de execucao

### Pelo menu

A funcao `MICMENU()` abre a tela de parametros e gera o arquivo no servidor.
O envio por e-mail e opcional.

| Parametro | Descricao | Padrao |
| --- | --- | --- |
| Emissao de | Data inicial do movimento | Data-base |
| Emissao ate | Data final do movimento | Data-base |
| Armazem de | Armazem inicial | Todos |
| Armazem ate | Armazem final | Todos |
| Enviar por e-mail | Envia o anexo apos a geracao | Nao |

### Pelo agendamento

A funcao `MICAUTO()` foi criada para o QBG. Ela:

1. Usa a data corrente como periodo.
2. Consulta todos os armazens.
3. Gera a planilha sem interface.
4. Envia o anexo automaticamente.

Quando executada fora de uma sessao Protheus, a rotina monta o ambiente da
empresa `03`, filial `5202`, e limpa apenas o ambiente que ela propria criou.

## Arquivo gerado

Os arquivos sao gravados em:

```text
\cachenosso\MIC_<identificador>.xlsx
```

A planilha possui uma aba chamada `MIC`, com as colunas:

`Filial`, `Armazem`, `Produto`, `Descricao`, `Lote`, `Tipo`, `Grupo`, `Und`,
`Dt Emissao`, `Centro Custo`, `Conta Contabil`, `Qtd Movto`, `Custo Unitario`,
`Custo Movto`, `Tipo Movto`, `Desc Tipo Movto`, `Documento`, `Sequencial`,
`Ordem Servico`, `Ordem Producao`, `Solicitante`, `Funcionario` e `Atendente`.

## Envio de e-mail

Os destinatarios principais sao configurados no parametro `ZZ_MOVCUST`, em
`SIGACFG > Ambiente > Cadastros > Parametros`, separados por ponto e virgula.
Assim, a lista pode ser alterada sem recompilar o fonte. O parametro e por
empresa e deve ser criado/configurado na empresa `03`.

O destinatario em copia continua definido no fonte [EST051.tlpp](EST051.tlpp):

```advpl
#DEFINE MIC_CC "ana.jaize@prolinhas.com.br"
```

O remetente e obtido pelo parametro `MV_RELACNT`. Se `ZZ_MOVCUST` ou
`MV_RELACNT` estiver vazio, o arquivo ainda pode ser gerado, mas o envio nao
sera realizado.

## Dependencias

- Protheus com suporte a TLPP.
- TopConn e acesso as tabelas `SD3`, `SB1`, `SCP`, `SRA`, `SF5` e `SYS_USR`.
- Classe `FWMsExcelXlsx` disponivel no ambiente.
- Funcao `u_WFEmail` configurada para envio automatico.
- Parametro `ZZ_MOVCUST` preenchido com os destinatarios principais.
- Parametro `MV_RELACNT` preenchido para execucoes com e-mail.

## Funcoes publicas

| Funcao | Uso |
| --- | --- |
| `RelatorioMIC()` | Consulta e gera o arquivo; pode enviar e-mail no modo automatico |
| `MICMENU()` | Interface de execucao manual |
| `MICAUTO()` | Execucao sem tela pelo agendamento |

Exemplo de chamada direta:

```advpl
// Modo manual, periodo e armazem informados
U_RelatorioMIC(.F., "20260825", "20260826", "AL", "AL")

// Modo automatico: usa a data corrente e envia por e-mail
U_RelatorioMIC(.T.)
```

## Observacoes tecnicas

- As datas recebidas por `RelatorioMIC()` usam o formato `AAAAMMDD`.
- Data em branco significa a data-base do Protheus.
- Armazem final em branco considera todos os armazens.
- O custo unitario fica zerado quando a quantidade do movimento e zero.
- Se nenhum registro for encontrado, nenhum arquivo ou e-mail e gerado.

## Historico

O relatorio consolida a consulta e as colunas do antigo `EST033` com a
execucao automatica que existia no `EST051`.

## SQL de referencia

Consulta usada como base para validar os dados do MIC. Os nomes das tabelas,
empresa, filial e periodo estao fixos neste exemplo; na rotina, esses valores
sao ajustados pelo ambiente e pelos parametros informados.

```sql
SELECT TOP 10
       D3_FILIAL,
       D3_LOCAL,
       D3_COD,
       D3_LOTECTL,
       D3_EMISSAO,
       D3_CC,
       D3_QUANT,
       D3_CUSTO1,
       D3_TM,
       D3_CF,
       D3_DOC,
       D3_NUMSEQ,
       D3_MAT,
       ISNULL(SRA.RA_NOME, '')  AS ranomefun,
       D3_USUARIO,
       D3_ORDEM,
       D3_OP,
       ISNULL(SRA2.RA_NOME, '') AS ranomesoli,
       B1_TIPO,
       B1_GRUPO,
       B1_UM,
       B1_CONTA,
       B1_DESC,
       ISNULL(CP_MAT, '')       AS cpmat,
       ISNULL(USR.USR_NOME, '') AS ranomeaten
  FROM SD3030 SD3
 INNER JOIN SB1030 SB1
    ON B1_FILIAL = '    '
   AND B1_COD = D3_COD
   AND SB1.D_E_L_E_T_ = ''
 INNER JOIN SF5010 SF5
    ON F5_FILIAL = '    '
   AND F4_TIPO <> 'P'
  LEFT JOIN SCP030 SCP
    ON CP_FILIAL = '5202'
   AND CP_NUM = D3_NUMSA
   AND CP_ITEM = D3_ITEMSA
   AND SCP.D_E_L_E_T_ = ''
  LEFT JOIN SRA030 SRA
    ON RA_FILIAL = '5202'
   AND SRA.RA_MAT = CP_MAT
   AND SRA.D_E_L_E_T_ = ''
  LEFT JOIN SRA030 SRA2
    ON SRA2.RA_FILIAL = '5202'
   AND SRA2.RA_MAT = D3_MAT
   AND SRA2.D_E_L_E_T_ = ''
  LEFT JOIN SYS_USR USR
    ON USR.USR_CODIGO = D3_USUARIO
   AND USR.D_E_L_E_T_ = ''
 WHERE SD3.D_E_L_E_T_ = ''
   AND D3_FILIAL = '5202'
   AND D3_ESTORNO <> 'S'
   AND D3_EMISSAO BETWEEN '20260825' AND '20260826'
   AND D3_LOCAL BETWEEN 'AL' AND 'AL'
   AND D3_CF <> 'DE4'
   AND D3_CF <> 'RE4';
```

> **Nota:** a implementacao atual corrige a ligacao com `SF5` usando
> `F5_CODIGO = D3_TM` e usa `F5_TEXTO` como descricao do tipo de movimento.