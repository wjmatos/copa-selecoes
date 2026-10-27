/* Dados do painel. Edite pelo admin.html ou direto aqui. */
const DATA={
 atualizado:"26/10/2026",

 grupo:"Grupo B · 39 a 46 anos · Masculino",
 votacao:{url:"",ativa:true,oculta:true},
 elencos:{
  "EUA":{r:"LUIZ FERNANDO SARRACINI",j:["ALVARO VINICIUS SANTOS HOGERA","ANDERSON JOSE LOPES FERREIRA","ANDRE LUIZ DE MIRANDA GALVÃO","DANILO MARTINS VIEIRA","EMILIO BARBOSA ALMEIDA","ESTEVAO PATERNO HILARIO","FULVIO DE BRITO BORGES","LUIS ANDRE DE MELO JUNIOR","LUIZ FERNANDO JERONIMO SILVA|G","LUIZ FERNANDO SARRACINI","PAULO HENRIQUE DOS SANTOS SILVA","RAFAEL ARAUJO DOS SANTOS","RAPHAEL ROMAO PENNA","RICARDO ANTONIO CABRAL","RICARDO INACIO DA SILVA","ROBSON DA SILVA COSTA","THIAGO DE OLIVEIRA SOUZA","THIAGO DOS SANTOS SILVA"]},
  "PERU":{r:"EDUARDO KARAKAMA SIQUEIRA",j:["ADRIANO ALVES DE OLIVEIRA DUARTE","ADRIANO MANOEL DA SILVA","ATANIBAL FERREIRA DE ABREU JUNIOR","BRUNNO NATALINO SENA","DIEGO DE SOUSA FREITAS","EDUARDO KARAKAMA SIQUEIRA","FERNANDO DE ARAUJO MARQUES","FERNANDO JOSE MALAGOGIN","JAQUES MAIA FERREIRA","KLEBER EVANDRO LOUREIRO","LEONARDO HENRIQUE SERAFIM DA SILVA","MARCEL DIEGO DANTAS MOURA","RODRIGO GONCALVES DOS SANTOS","RUBENS GONZAGA DOS SANTOS|G","THIAGO MACEDO PEREIRA DE MATOS","VITOR AUGUSTO GONÇALVES","WELLINGTON DA CUNHA PEREIRA","WILOMARK SANTOS DE OLIVEIRA"]},
  "URUGUAI":{r:"LUCAS RODRIGUES DA SILVA",j:["ANTONIO ELIELTON ALVES DO NASCIMENTO","CAETANO DO NASCIMENTO MOTOSHIMA","CELSO ABREU PEREIRA","DANILO SUGIMOTO DE ARAUJO","EDUARDO FERRAZ COUTO","FLAVIO JOSE VANNUCCI DE MAGALHAES","LUCAS RODRIGUES DA SILVA","LUCIANO CORDEIRO DA ROCHA","MARCELO DA SILVA CARDOSO","PAULO SERGIO DE JESUS SANTOS|G","PETERSON TEIXEIRA DARIO","RAPHAEL ARAUJO LEITE","RENATO DE OLIVEIRA MELO","RICARDO SENA MACHADO","ROBSON EDUARDO CAMERA","RODRIGO DE SOUZA SOARES","SAULO OLIVEIRA ALONSO DE REZENDE","WELLINGTON TIAGO PINHEIRO"]},
  "MÉXICO":{r:"MARCELO DA SILVA MORAIS",j:["ANDERSON FREIRE DOS SANTOS","ANDRE ALMEIDA DOS SANTOS","ANDRE DE LUCENA FEITOSA","ANDRE LUIZ CELESTINO","ANTONIO CARLOS LOCONTE DA SILVA","CARLOS HENRIQUE DA SILVA SANTOS","CELINO JOSE OLIVEIRA NETO","EVERTON SIMS DE QUEIROZ","FABIO MARTINS DA COSTA","JOAO PAULO DOS REIS NERY","JOSIAS ALVES DA SILVA","KLEBER ROGERIO MARTINS","MARCELO DA SILVA MORAIS","RAFAEL MARQUES DE OLIVEIRA","ROMARIO FERNANDES DA CRUZ","THIAGO BARRICHELLO SIGALI","VALDEMIR JOAO GOMES","WAGNER DE CARVALHO|G"]},
  "ARGENTINA":{r:"MARCEL PAULINO",j:["ALEX DE ANDRADE","ANDERSON SOUZA SANTOS","ANDRE LUIZ DE JESUS MARQUES|G","CAIO CÉSAR DE ALMEIDA LIMA","CESAR CONSOLARO MENDONCA","FERNANDO BRONZATI LEMOS","JOSÉ WANDERSON G DA SILVA","LEANDRO RODRIGUES TORRECILHAS","MARCEL PAULINO","MARCELO ESUTACHIO DOS SANTOS SOARES","MARCELO GONZALEZ ROCHA","MARCIO ANTONIO DOS ANJOS","REINALDO SOUZA DOS SANTOS","ROBERTO CIARCIA JUNIOR","ROBERTO LUIS HOFF PALMARI","THOMAS LIRA DE SOUSA","WAGNER RODRIGUES COSTA","WELLINGTON NOGUEIRA CAVALCANTE"]},
  "CHILE":{r:"THIAGO BAUER",j:["ANDRE ATAIDE DE LIMA","CARLOS EDUARDO DOS SANTOS SOUZA|G","CLAYTON COSTA GONÇALVES DIAS","DIEGO GOMES DA SILVA ALVES","EDUARDO DE LIMA PINTO","EDUARDO SANTO","FLAVIO ALVES DA SILVA","HELMER MACEDO DANTAS","JOAO CARLOS BONFIM TEIXEIRA","MARCOS VINICIUS SOUTO DAMASIO","MICHAEL RODRIGUES DE MOURA","MIKE ALVES GOMES","PAULO DE JESUS SANTOS","RAPHAEL FLORES MOURA","RENATO DE ABREU CAVALCANTE","THIAGO BAUER","TIAGO PEREIRA CAZZOTO","WAGNER EDUARDO SOUSA DOS SANTOS","WILLYAN MAGALHAES DE OLIVEIRA"]},
  "PARAGUAI":{r:"DENIS PEREIRA DE ALMEIDA",j:["DANILO AUGUSTO PEREIRA BENASSI","BRUNO COELHO HONDA|G","DAVI RODRIGUES MONTALVAO","DENIS PEREIRA DE ALMEIDA","EDUARDO LUIZ DE ASSIS BRAZIL","FRANCISCO DO CARMO MONTSERRAT","GUILHERME SILVA DE JESUS","JOSE PEREIRA LOPES","LEONARDO LIMA GONZAGA DE SOUZA","MARCIO DE SOUZA COELHO","MARCOS FERREIRA DE MIRANDA","MIQUEIAS SILVA PEIXOTO","RAFAEL FERREIRA DE MIRANDA","RENATO ROGENSKI","RICARDO EDSON DE SOUZA","ROGERIO ALVES GARCIA","RONALDO CARVALHO DE SOUSA","VANTUIL JOSE PIRES PEDRA","VICTOR SILVA BELLINI"]},
  "BRASIL":{r:"WILLIAN APARECIDO MATOS",j:["ANDERSON BORGES PEREIRA","AUGUSTO CESAR CAMPANILE","DAVI FREITAS DOS SANTOS|G","DIEGO AUGUSTO DE MATTOS","FABIANO OLIVEIRA ARAUJO","FRANCENILTON DA SILVA DIAS","JULIO CESAR SILVA","LEANDRO PACHECO GONÇALVES","LUIS FERNANDO FERREIRA BORGES","LUIZ HENRIQUE OLIVEIRA FERREIRA","MARCELO WILLIAM MARTINS","MARCOS ANDRE DANTAS DA CRUZ","RAFAEL ASSENOFF ALMEIDA","RODRIGO FELIPE DE LIMA","THIAGO DE OLIVEIRA MOURA","VINICIUS SOARES FERREIRA","WILLIAN APARECIDO MATOS","YURI FIGUEREDO SPACOV"]}
 },
 times:["EUA","PERU","URUGUAI","MÉXICO","ARGENTINA","CHILE","PARAGUAI","BRASIL"],
 cor:{EUA:"#1f3a93",PERU:"#a50f2a",URUGUAI:"#1c5aa6","MÉXICO":"#1b8a3a",ARGENTINA:"#4aa3dd",CHILE:"#d4421f",PARAGUAI:"#8e2457",BRASIL:"#c99700"},
 sig:{EUA:"EUA",PERU:"PER",URUGUAI:"URU","MÉXICO":"MEX",ARGENTINA:"ARG",CHILE:"CHI",PARAGUAI:"PAR",BRASIL:"BRA"},
 jogos:[
  {n:1,data:"2026-09-12",hora:"13:10",casa:"ARGENTINA",fora:"URUGUAI",placar:[1,2]},
  {n:2,data:"2026-09-12",hora:"14:50",casa:"CHILE",fora:"PARAGUAI",placar:[1,1]},
  {n:3,data:"2026-09-12",hora:"16:30",casa:"BRASIL",fora:"PERU",placar:[0,3]},
  {n:4,data:"2026-09-13",hora:"13:10",casa:"EUA",fora:"MÉXICO",placar:[3,2]},
  {n:6,data:"2026-09-17",hora:"20:00",casa:"PERU",fora:"CHILE",placar:[1,1]},
  {n:7,data:"2026-09-20",hora:"13:10",casa:"ARGENTINA",fora:"BRASIL",placar:[1,0]},
  {n:8,data:"2026-09-22",hora:"20:00",casa:"PARAGUAI",fora:"EUA",placar:[3,2]},
  {n:5,data:"2026-09-23",hora:"20:00",casa:"URUGUAI",fora:"MÉXICO",placar:[2,6]},
  {n:11,data:"2026-09-27",hora:"13:10",casa:"BRASIL",fora:"URUGUAI",placar:[1,5]},
  {n:12,data:"2026-09-27",hora:"14:50",casa:"EUA",fora:"PERU",placar:[1,5]},
  {n:13,data:"2026-09-29",hora:"20:00",casa:"MÉXICO",fora:"CHILE",placar:[2,0]},
  {n:14,data:"2026-10-03",hora:"09:50",casa:"BRASIL",fora:"MÉXICO",placar:[2,1]},
  {n:15,data:"2026-10-03",hora:"11:30",casa:"ARGENTINA",fora:"PARAGUAI",placar:[3,2]},
  {n:16,data:"2026-10-03",hora:"13:10",casa:"URUGUAI",fora:"PERU",placar:[0,0]},
  {n:17,data:"2026-10-03",hora:"14:50",casa:"CHILE",fora:"EUA",placar:[0,5]},
  {n:18,data:"2026-10-06",hora:"20:00",casa:"EUA",fora:"URUGUAI",placar:[2,1]},
  {n:9,data:"2026-10-13",hora:"20:00",casa:"CHILE",fora:"ARGENTINA",placar:[4,1]},
  {n:20,data:"2026-10-16",hora:"20:00",casa:"ARGENTINA",fora:"EUA",placar:[3,5]},
  {n:19,data:"2026-10-17",hora:"13:10",casa:"PERU",fora:"MÉXICO",placar:[1,1]},
  {n:21,data:"2026-10-18",hora:"13:10",casa:"URUGUAI",fora:"PARAGUAI",placar:[1,3]},
  {n:22,data:"2026-10-18",hora:"14:50",casa:"BRASIL",fora:"CHILE",placar:[2,1]},
  {n:23,data:"2026-10-20",hora:"20:00",casa:"PERU",fora:"ARGENTINA",placar:[1,2]},
  {n:28,data:"2026-10-22",hora:"20:00",casa:"PARAGUAI",fora:"BRASIL",placar:[1,1]},
  {n:24,data:"2026-10-25",hora:"11:00",casa:"CHILE",fora:"URUGUAI",placar:[3,0]},
  {n:25,data:"2026-10-25",hora:"12:40",casa:"EUA",fora:"BRASIL",placar:[1,2]},
  {n:26,data:"2026-10-25",hora:"14:20",casa:"MÉXICO",fora:"ARGENTINA",placar:[3,0]},
  {n:27,data:"2026-10-25",hora:"16:00",casa:"PARAGUAI",fora:"PERU",placar:[3,0]},
  {n:10,data:"2026-10-31",hora:"09:00",casa:"MÉXICO",fora:"PARAGUAI",placar:null}
 ],
 artilharia:[],
 cartoes:[],
 noticias:[
  {tag:"Domingo, 25/10",t:"Brasil vence os EUA por 2 a 1 e chega a 10 pontos",x:"O Brasil sobe para o 4º lugar e já encerrou seus jogos na fase de grupos. Os EUA seguem com 12 pontos e também terminaram suas partidas.",times:["BRASIL","EUA"]},
  {tag:"Domingo, 25/10",t:"Paraguai goleia o Peru por 3 a 0 e assume o 2º lugar",x:"Com 11 pontos e um jogo a menos, o Paraguai fica a 1 ponto dos EUA. O Peru, que perdeu duas seguidas, cai para o 6º lugar com 9 pontos.",times:["PARAGUAI","PERU"]},
  {tag:"Último jogo",t:"México x Paraguai, em 31/10, define o 1º lugar",x:"É a única partida que falta na fase de grupos. Quem vencer pode terminar na ponta: o Paraguai (11 pontos) chega a 14 e o México (10) a 13. Em caso de empate, os EUA (12) ficam à frente pelo número de vitórias.",times:["MÉXICO","PARAGUAI"]},
  {tag:"Domingo, 25/10",t:"México goleia a Argentina por 3 a 0",x:"O México chega a 10 pontos e fica em 3º, com um jogo a menos. A Argentina cai para 9 pontos e para o 5º lugar.",times:["MÉXICO","ARGENTINA"]},
  {tag:"Domingo, 25/10",t:"Chile vence o Uruguai por 3 a 0 e deixa a lanterna",x:"O Chile chega a 8 pontos e sobe para o 7º lugar. O Uruguai termina com 7 pontos, é o último colocado e, pelo formato informado, está eliminado.",times:["CHILE","URUGUAI"]},
  {tag:"Quinta, 22/10",t:"Paraguai e Brasil empatam em 1 a 1",x:"O Paraguai chega a 8 pontos e o Brasil soma 7.",times:["PARAGUAI","BRASIL"]},
  {tag:"Terça, 20/10",t:"Argentina vence o Peru por 2 a 1 e chega a 9 pontos",x:"O Peru perde a invencibilidade e a Argentina chega a 9 pontos.",times:["ARGENTINA","PERU"]},
  {tag:"Defesa",t:"Peru tem a defesa menos vazada pela média",x:"8 gols sofridos em 7 jogos (1,1 por partida). Paraguai e México também têm 8, mas em 6 jogos (1,3).",times:["PERU"]},
  {tag:"Ataque",t:"EUA têm o melhor ataque: 19 gols em 7 jogos",x:"Média de 2,7 gols por partida. O México vem em seguida, com 2,5 por jogo (15 em 6 jogos).",times:["EUA","MÉXICO"]}
 ],
 notas:[
  {tag:"Regra em foco",t:"3 amarelos = suspensão no jogo seguinte",x:"Art. 29: quem acumula 3 amarelos fica fora da próxima partida. Ao fim da 1ª fase, quem tem até 2 amarelos é anistiado."},
  {tag:"Regra em foco",t:"20 minutos obrigatórios para todos da súmula",x:"Art. 19: todo jogador relacionado precisa jogar no mínimo 20 minutos (15 no torneio início)."}
 ],
 brechas:[]
};
