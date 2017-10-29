################################################
##### PROGRAMADO POR: LEANDRO L. MONTANARI #####
################################################

from win32com.client import Dispatch as voz

falar = voz('SAPI.SpVoice')
som = ''
tutorial = True

while som != 'S' and som != 'N' and som != 'SIM' and som != 'NÃO' and som != 'NAO':
    som = str(input('Habilitar narração [S/N]? ')).strip().upper().replace(' ','')
    print('')

if som == 'SIM':
    som = 'S'
elif som == 'NÃO':
    som = 'N'
elif som == 'NAO':
    som = 'N'

d = {
    3 : ['MIL', 'MIL'],
    6 : ['MILHÃO', 'MILHÕES'],
    9 : ['BILHÃO', 'BILHÕES'],
    12 : ['TRILHÃO', 'TRILHÕES'],
    15 : ['QUATRILHÃO', 'QUATRILHÕES'],
    18 : ['QUINTILHÃO', 'QUINTILHÕES'],
    21 : ['SEXTILHÃO', 'SEXTILHÕES'], 
    24 : ['SEPTILHÃO', 'SEPTILHÕES'],
    27 : ['OCTILHÃO', 'OCTILHÕES'],
    30 : ['NONILHÃO', 'NONILHÕES'],
    33 : ['DECILHÃO', 'DECILHÕES'],
    36 : ['UNDECILHÃO', 'UNDECILHÕES'],
    39 : ['DUODECILHÃO', 'DUODECILHÕES'],
    42 : ['TREDECILHÃO', 'TREDECILHÕES'],
    45 : ['QUATTUORDECILHÃO', 'QUATTUORDECILHÕES'],
    48 : ['QUINDECILHÃO', 'QUINDECILHÕES'],
    51 : ['SEXDECILHÃO', 'SEXDECILHÕES'],
    54 : ['SEPTENDECILHÃO', 'SEPTENDECILHÕES'],
    57 : ['OCTODECILHÃO', 'OCTODECILHÕES'],
    60 : ['NOVENDECILHÃO', 'NOVENDECILHÕES'],
    63 : ['VIGINTILHÃO', 'VIGINTILHÕES'],
    66 : ['UNVIGINTILHÃO', 'UNVIGINTILHÕES'],
    69 : ['DUOVIGINTILHÃO', 'DUOVIGINTILHÕES'],
    72 : ['TREVIGINTILHÃO', 'TREVIGINTILHÕES'],
    75 : ['QUATTUORVIGINTILHÃO', 'QUATTUORVIGINTILHÕES'],
    78 : ['QUINVIGINTILHÃO', 'QUINVIGINTILHÕES'],
    81 : ['SEXVIGINTILHÃO', 'SEXVIGINTILHÕES'],
    84 : ['SEPTENVIGINTILHÃO', 'SEPTENVIGINTILHÕES'],
    87 : ['OCTOVIGINTILHÕES', 'OCTOVIGINTILHÕES'],
    90 : ['NOVENVIGINTILHÃO', 'NOVENVIGINTILHÕES'],
    93 : ['TRIGINTILHÃO', 'TRIGINTILHÕES'],
    96 : ['UNTRIGINTILHÃO', 'UNTRIGINTILHÕES'],
    99 : ['DUOTRIGINTILHÃO', 'DUOTRIGINTILHÕES'],
    102 : ['TRETRIGINTILHÃO', 'TRETRIGINTILHÕES'],
    105 : ['QUATTUORTRIGINTILHÃO', 'QUATTUORTRIGINTILHÕES'],
    108 : ['QUINTRIGINTILHÃO', 'QUINTRIGINTILHÕES'],
    111 : ['SEXTRIGINTILHÃO', 'SEXTRIGINTILHÕES'],
    114 : ['SEPTENTRIGINTILHÃO', 'SEPTENTRIGINTILHÕES'],
    117 : ['OCTOTRIGINTILHÃO', 'OCTOTRIGINTILHÕES'],
    120 : ['NOVENTRIGINTILHÃO', 'NOVENTRIGINTILHÕES'],
    123 : ['QUADRAGINTILHÃO', 'QUADRAGINTILHÕES']
    }

final = sorted(d.keys())[-1]
print('Atualmente, o programa suporta números com até {} zeros.\n'.format(final + 2))

while True:
    try:
        if tutorial:
            msg = 'Digite um número gigante para conhecer seu nome (ex: 1000000000000000000000000000000000000000): '
        else:
            msg = 'Digite um número gigante para conhecer seu nome: '

        n = int(input(msg))
        sn = str(n)
        tam = len(sn)
        tutorial = False

        for i in range(3, final + 1, 3):  # Repete o laço para verificar a quantidade de zeros ou dígitos
            if n >= 10 ** i and n < 10 ** (i + 3):  # Verifica se o número está na faixa de i a (i + 3)               
                if sn[0] == '1' and tam == i + 1:  # Verifica se é singular
                    frase = '\n1 {}\n'.format(d[i][0])
                    print(frase)
                    if som == 'S':  # Se o som estiver habilitado, além de exibir na tela, também fala o número
                        falar.Speak(frase)
                else:  # Ou plural
                    frase = '\n{} {}\n'.format(sn[0:tam - i], d[i][1])
                    print(frase)
                    if som == 'S':  # Se o som estiver habilitado, além de exibir na tela, também fala o número
                        falar.Speak(frase)
                break
            else:
                if i == final:  # Se nenhuma condição for suprida até o final do laço, então exibe apenas o próprio número
                    print('\n{}\n'.format(n))
            
    except ValueError:
        print('\nValor inválido. Digite um número!\n')
        continue
