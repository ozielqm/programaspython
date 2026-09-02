# Beat & Repeat

App de inglês para celular que ensina pelo **ritmo**. Nível A1–A2, interface em português.

O arquivo `chant.html` é o app inteiro: um único HTML, sem servidor, sem instalação,
sem dependências. Abra num navegador e funciona.

## O método

Cada frase é um **jazz chant** (método da Carolyn Graham, usado em sala de aula desde os
anos 1970): a frase é encaixada num compasso de 4 tempos e a **sílaba tônica cai na batida**.

O aluno brasileiro costuma travar na fala rápida não por falta de vocabulário, mas porque o
inglês é uma língua de ritmo acentual — as sílabas fracas encolhem entre as fortes. Material
tradicional ensina as palavras e não ensina esse encaixe. O chant ensina, porque o corpo pega
o ritmo antes de a cabeça traduzir.

Não há música comercial no app: nada de Spotify, YouTube ou MP3 de artista. Isso é decisão de
projeto, não limitação disfarçada — conteúdo de terceiros não pode ser distribuído junto. A
batida é sintetizada com Web Audio e a voz vem do sintetizador do próprio aparelho, então todo
o conteúdo é original e nada depende de rede depois que a página abre.

## Como abrir no celular

1. Abra o link do app no navegador do celular (Chrome no Android, Safari no iPhone).
2. Menu do navegador → **Adicionar à tela de início**.
3. Abra pelo ícone. Ele roda em tela cheia, sem barra de navegador.

Para usar sem link, é só copiar `chant.html` para o aparelho e abrir pelo gerenciador de
arquivos — o app é autossuficiente.

## Como usar

Escolha um chant e toque no botão vermelho. Cada chant tem 4 etapas, e a seguinte só abre
depois da anterior:

| Etapa | O que acontece |
|---|---|
| 1. Ouvir | O app toca a batida e fala as linhas; as sílabas acendem no tempo. |
| 2. Repetir | O app fala uma linha e deixa um compasso inteiro em silêncio para você repetir. |
| 3. Acompanhar | Só batida e palavras acendendo. Você conduz. |
| 4. Acelerar | O mesmo chant na escada 76 → 88 → 100 → 112 BPM. |

Controles no rodapé: play/pause, andamento (BPM), voz liga/desliga e tradução mostra/esconde.
O progresso e a sequência de dias ficam salvos no `localStorage` **daquele aparelho** — trocar
de celular ou limpar os dados do navegador zera o progresso.

## Conteúdo

10 chants originais cobrindo o programa A1–A2: cumprimentos, verbo *to be*, perguntas com
*wh-*, números, família, rotina e horas, *like / don't like*, *can / can't*,
*there is / there are* e *present continuous*. São 60 linhas, todas com tradução.

## Como acrescentar um chant

O ritmo é escrito direto no dado, no array `CHANTS`. A notação é:

- `|` separa as 4 batidas do compasso
- espaço separa palavras dentro da mesma batida
- `-` separa sílabas
- o **primeiro token de cada batida é o que cai na batida** (a tônica)
- batida vazia = pausa, e a pausa é parte do balanço — use sem medo

```js
L("Hel-lo, | hel-lo, | how are | you?", "Olá, olá, como vai você?")
//  ↑ Hel        ↑ hel      ↑ how      ↑ you    ← as quatro batidas
```

Uma palavra pode atravessar a barra de compasso terminando em hífen:
`L("No, there | is-n't. | Look a- | gain!", ...)` é falado como *"No, there isn't. Look again!"*.

O texto que a voz fala é derivado do próprio padrão, então não existe risco de a letra escrita
e a letra falada saírem diferentes.

## Limites conhecidos

- **Offline no arranque a frio não é garantido.** Um *service worker* precisaria de um arquivo
  `.js` separado, o que não existe num app de arquivo único. Na prática o app não faz nenhuma
  chamada de rede depois de carregado e o cache normal do navegador cobre o uso diário — mas
  desinstalar/limpar o cache exige estar online uma vez para reabrir.
- **A voz depende do aparelho.** O app usa `speechSynthesis` com as vozes instaladas no celular.
  Sem voz em inglês, ele avisa na tela e as etapas *Acompanhar* e *Acelerar* seguem funcionando.
- **A voz não canta e não é milimetricamente sincronizada** com a batida — quem sustenta o ritmo
  é a batida e o destaque visual, que é como um jazz chant funciona em sala também.
- **Sem microfone nesta versão.** Na etapa *Acelerar* quem confirma se acompanhou é você. Dá
  para acrescentar avaliação de pronúncia por reconhecimento de voz depois.
- **As fontes (Archivo e DM Mono) vêm do Google Fonts.** Sem internet elas caem para as fontes
  do sistema e o app continua idêntico em funcionamento — muda só o desenho das letras.

## Verificação

Testado com Playwright/Chromium em viewport de celular (390×844): carga sem erros de script,
`AudioContext` destravando no toque, batidas saindo no intervalo certo a 76 e 112 BPM
(783 ms e 533 ms medianos, contra 789 ms e 536 ms teóricos), ciclo 1‑2‑3‑4 correto, sílaba
acendendo na batida certa, progresso gravando e relendo do `localStorage`, e as 60 linhas com
4 batidas, tradução e texto falado limpo.
