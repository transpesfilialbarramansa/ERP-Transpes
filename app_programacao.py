import io
import base64
import datetime
import hashlib
import json
import psycopg2
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu

# Importações para geração de PDF e Excel
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Transpes - Sistema de Gestão",
    page_icon="logo_transpes.png",
    layout="wide"
)

# ==========================================
# CONEXÃO E BANCO DE DADOS (SUPABASE / POSTGRESQL)
# ==========================================

def get_connection():
    return psycopg2.connect(st.secrets["DB_URL"])

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# DADOS INICIAIS DE FORNECEDORES EXTRAÍDOS DA PLANILHA
FORNECEDORES_INICIAIS = [
    ('99', 'ALEXANDRE SILVANO DE MELO', 'ALEMAO MUNCK', 'CAETITE', 'BA', '28.332.244/0001-90', '77 9925-3952'),
    ('108', 'TERMOSOL CONSTRUTORA E COMERCIO LTDA', 'CONSTRUTORA TERMOSOL', 'CAETITE', 'BA', '06.872.066/0001-58', '77 3454-3659'),
    ('43', 'AGM CONSTRUTORA LTDA', 'AGM CONSTRUTORA', 'CALMON', 'BA', '11.051.592/0001-97', '74 36212416'),
    ('151', 'PEZINHO GUINCHO LOCACAO E TRANSPORTE LTDA - ME', 'PEZINHO GUINCHO LOCACAO E TRANSPORTE', 'FEIRA DE SANTANA', 'BA', '13.332.190/0001-96', '75 36234413'),
    ('148', 'GUINCHO GRAPIUNA SERVICOS E COMERCIO LTDA', 'GUINCHO GRAPIUNA SERVICOS E COMERCIO LTDA', 'ITABUNA', 'BA', '15.179.807/0001-00', '73 2117-567'),
    ('138', 'TRANSDULTRA TRANSPORTES E SERVICOS LTDA', 'TRANSDULTRA', 'SALVADOR', 'BA', '34.031.997/0001-69', '71 3392-3783'),
    ('103', 'MARCELO JESUS DA SILVA', 'STOPCAR GUINCHO', 'SEABRA', 'BA', '09.311.489/0001-97', '75 33312219'),
    ('2614', 'FRETEBRAS INTERNET E SERVIÇOS LTDA', 'FRETEBRAS', 'NACIONAL', 'BR', '10.885.840/0002-13', '64 34425221'),
    ('48', 'RAIMUNDO NONATO ALVES CAMELO', 'RAIMUNDO NONATO ALVES CAMELO', 'CRATEUS', 'CE', '147.886.208-47', '88 9983-0244'),
    ('129', 'J L RODRIGUES TRANSPORTE - ME', 'LOG MUNCK SERVICOS', 'FORTALEZA - IRAUCUBA - HORIZONTE - BARREIRA - MARACANAU', 'CE', '12.407.910/0001-72', '85 86650251'),
    ('192', 'J L RODRIGUES TRANSPORTE - ME', 'LOG MUNCK SERVICOS', 'FORTALEZA - IRAUCUBA - HORIZONTE - BARREIRA - MARACANAU', 'CE', '12.407.910/0001-72', '85 86650251'),
    ('139', 'BOTO AUTO TRUCK LTDA', 'BOTO AUTO TRUCK', 'TIANGUA', 'CE', '24.389.791/0001-51', '88 36711600'),
    ('180', 'ECAZ TRANSPORTES E SERVICOS LTDA.', 'ECAZ TRANSPORTES E SERVICOS LTDA', 'BRASILIA', 'DF', '04.866.758/0001-68', '61 996080944'),
    ('187', 'REAL LOCADORA DE CAMINHOES E EQUIPAMENTOS LTDA', 'REAL LOCADORA', 'BRASILIA', 'DF', '05.278.358/0001-02', ''),
    ('115', 'SOYARA DE FATIMA LOPES MADEIRA', 'SOYARA DE FATIMA LOPES MADEIRA', 'CARIACICA', 'ES', '818.464.877-49', '27 999820237'),
    ('71', 'DOMINGOS DE ALMEIDA SANTOS', 'DOMINGOS MUNCK', 'CARIACICA - SERRAS', 'ES', '11.196.587/0001-72', '27 99518675'),
    ('146', 'J M CORDEIRO MANUTENCOES ELETRICAS LTDA', 'J M CORDEIRO MANUTENCOES ELETRICAS', 'GUARAPARI', 'ES', '03.659.553/0001-49', ''),
    ('112', 'BERNARDO JECKEL JUNIOR', 'BERNARDO JECKEL JUNIOR', 'SANTA MARIA DE JETIBA', 'ES', '083.380.257-71', '27 99966-7588'),
    ('179', 'R & R LOCACAO & SERVICOS LTDA', 'RAMOS & RAMOS', 'SANTA PAULA - VILA VELHA', 'ES', '06.038.545/0001-73', '27 32442158'),
    ('170', 'PL LOCACOES E TRANSPORTES EIRELI ME', 'PL LOCACOES E TRANSPORTES EIRELI ME', 'SERRA', 'ES', '26.590.250/0001-12', '27 30571567'),
    ('100', 'EDITE DE SOUZA PEDRO - EPP', 'EDITE DE SOUZA PEDRO - EPP', 'SERRA - VILA VELHA - CARIACICA - VITORIA', 'ES', '05.724.475/0001-44', '27 32389756'),
    ('85', 'GOIANIA LOCACOES DE GUINDASTES UNIPESSOAL LIMITADA', 'GOIANIA MUNCK', 'APARECIDA DE GOIANIA', 'GO', '30.751.456/0001-54', '62 40184040'),
    ('132', 'CESAR TRANSPORTES, GUINDASTES E EQUIPAMENTOS LTDA', 'CESAR TRANSPORTES, GUINDASTES E EQUIPAMENTOS LTDA', 'APARECIDA DE GOIANIA', 'GO', '00.148.726/0003-38', '62 32649500'),
    ('121', 'MARCELO LOURENCO BORGES 56539517149', 'MARCELO LOURENCO BORGES 56539517149', 'GOIANIA', 'GO', '31.293.283/0001-30', '62 9616-0002'),
    ('144', 'SAMELA AVELINO DOS SANTOS 01053240198', 'C&E TRANSPORTES E GUINDASTES', 'GOIANIA', 'GP', '42.997.945/0001-70', '62 8316-8856'),
    ('37', 'DARA CRISTINA INOCENCIO 14440291694', 'PEDRO TRANSPORTES', 'ARAXA', 'MG', '24.944.345/0001-61', '34 8881-9672'),
    ('107', 'TERRATRAN TERRAPLENAGEM E TRANSPORTES LTDA', 'TERRATRAN', 'BOTELHOS', 'MG', '03.104.452/0001-01', '35 3571-1388'),
    ('54', 'REUBERT CIMINI ME', 'SERRALHERIA DO RUI', 'CARATINGA - INHAPIM', 'MG', '25.401.771/0001-11', '33 99974757'),
    ('36', 'TERRAPLENAGEM SOUZA E FILHOS LTDA', 'SOUZA E FILHOS', 'CATAGUASES', 'MG', '23.245.400/0001-62', '32 3421-5156'),
    ('106', 'VALDECI DE OLIVEIRA REZENDE', 'VALDECI DE OLIVEIRA REZENDE', 'CATAGUASES', 'MG', '462.993.866-49', '32 98422-0060'),
    ('26', 'SOCORRO E TRANSPORTE DIAS EIRELI', 'SOCORRO E TRANSPORTE DIAS', 'CORONEL FABRICIANO', 'MG', '36.208.232/0001-87', '31 3846-0683'),
    ('49', 'L.S.A TRANSPORTES LTDA', 'L.S.A TRANSPORTES E LOCACOES', 'CORONEL FABRICIANO', 'MG', '05.233.958/0001-46', '31 3826-4445'),
    ('51', 'SIMAO PEDRO DE SOUZA', 'SIMAO PEDRO DE SOUZA', 'DIAMANTINA', 'MG', '13.211.308/0001-28', '38 35318844'),
    ('67', 'CARLOS ANTONIO BEZERRA SOARES CPF: 331.016.204-49', 'CARLINHOS GUINDAUTO', 'DIVINOPOLIS', 'MG', '06.302.898/0001-39', '37 32141552'),
    ('109', 'VITORIO LOCACOES - EIRELI', 'VITORIO LOCACOES', 'JUIZ DE FORA', 'MG', '35.453.457/0001-36', '32 9120-9797'),
    ('110', 'RENTMAQ LTDA', 'RENTMAQ LTDA', 'JUIZ DE FORA', 'MG', '71.259.097/0001-08', '32 3215-8108'),
    ('186', 'EZEQUIEL GUINCHOS E GUINDASTES LTDA', 'EZEQUIEL GUINCHOS E GUINDASTES', 'JUIZ DE FORA', 'MG', '32.965.718/0001-09', ''),
    ('78', 'SILVIO MARQUES LOUSADA JUNIOR E CIA LTDA', 'AUTO SOCORRO REBOCAR', 'LAVRAS - CARMO DA CACHOEIRA', 'MG', '13.336.911/0001-36', '35 3822-7405'),
    ('73', 'GERALDO NOGUEIRA PEREIRA', 'CERAMICA MINA NOVA', 'MINAS NOVAS', 'MG', '16.913.907/0001-81', '33 37641116'),
    ('96', 'MG GUINDASTES E LOCACAO DE EQUIPAMENTOS LTDA', 'MG GUINDASTES E LOCACAO DE EQUIPAMENTOS LTDA', 'PARACATU', 'MG', '12.464.432/0001-32', '38 3671-1985'),
    ('24', 'ED MUNCK SERVICOS E TRANSPORTES LTDA', 'ED MUNCK SERVICOS E LOCACOES', 'PASSOS', 'MG', '30.763.234/0001-51', '35 3526-0244'),
    ('79', 'TRAMEL TRANSPORTES E SERVICOS DE MUNCK EIRELI', 'TRAMEL TRANSPORTES E SERVICOS DE MUNCK EIRELI', 'PATOS DE MINAS', 'MG', '34.727.926/0001-03', '34 3823-8615'),
    ('94', 'GUINDASTES MARABA LTDA', 'GUINDASTES MARABA LTDA', 'PATOS DE MINAS - RIO PARANAIBA', 'MG', '19.716.390/0001-29', '34 38216688'),
    ('83', 'LEVI MENDES PEREIRA DOS SANTOS', 'AUTO SOCORRO PATROCINIO', 'PATROCINIO', 'MG', '27.425.267/0001-57', '34 3831-5230'),
    ('117', 'TRANSMEDEIROS GUINDASTES LTDA', 'TRANSMEDEIROS GUINDASTES', 'POCOS DE CALDAS', 'MG', '10.334.821/0001-38', '35 37142750'),
    ('84', 'R L LOCACAO DE MÁQUINAS E EOUIPAMENTOS EIRELI', 'R L LOCACAO DE MAQUINAS E EQUIPAMENTOS', 'POUSO ALEGRE', 'MG', '27.817.828/0001-49', '35 3422-5444'),
    ('143', 'GUINDASTES ALVINOPOLIS EIRELI', 'MARINHO GUINDASTES E TRANSPORTES', 'RIO PIRACICABA - ALVINOPOLIS - OURO PRETO', 'MG', '42.924.316/0001-90', '31 38551105'),
    ('30', 'PAULO CESAR BARBOSA DA SILVA', 'SOS GUINCHO SANTA BARBARA', 'SANTA BARBARA - BARÃO DE COCAIS', 'MG', '18.423.477/0001-30', '31 3832-3498'),
    ('102', 'L. A. F. DE SOUZA AUTO SOCORRO ME', 'L. A. F. DE SOUZA AUTO SOCORRO ME', 'SÃO GOTARDO - RIO PARANAIBA - MATUTINA - GUIMARANIA', 'MG', '08.683.842/0001-90', '34 36711585'),
    ('118', 'MARCOS ANTONIO CARDOSO 24195171691', 'CARDOSO TRANSPORTES E GUINCHO', 'SETE LAGOAS - INIMUTABA', 'MG', '32.148.971/0001-32', '31 37722748'),
    ('184', 'EXPRESSO SAO LUIZ DE MANHUACU LTDA', 'TRANSLUIZ', 'UBERABA', 'MG', '22.259.600/0001-44', '34 33181800'),
    ('29', 'CONTRAN CONSTRUTORA E TRANSPORTE LTDA', 'CONTRAN', 'UBERLANDIA', 'MG', '17.821.921/0001-48', '34 3219-0118'),
    ('18', 'A COSTA SERVICOS DE GUINCHO E CARRETO ME', 'AUTO SOCORRO COSTA', 'VARGINHA - MONSENHOR PAULO', 'MG', '03.883.336/0001-09', '35 3221-5080'),
    ('104', 'S. E. M. T. REMOÇÕES LTDA', 'LOGTRANS REMOCOES', 'CAMPO GRANDE', 'MS', '32.748.163/0001-81', '67 3388-0050'),
    ('131', 'VANDERLEI BENEDITO DE OLIVEIRA 19893910100', 'VANDERLEI BENEDITO DE OLIVEIRA', 'CUIABA - CACERES', 'MT', '36.853.518/0001-04', '65 36245366'),
    ('14', 'LOGMUNCK COMERCIO DE EQUIPAMENTOS E SERVICOS EIRELI', 'LOGMUNCK', 'CUIABA - RONDONOPOLIS', 'MT', '09.309.433/0001-81', '65 3682-1000'),
    ('135', 'J R P DE QUEIROZ', 'JR GUINCHO E MUNCK', 'LUCAS DO RIO VERDE - SORRISO - MUTUM - TAPURAH - NOVA MUTUM', 'MT', '20.730.015/0001-00', '65 35492476'),
    ('150', 'DISA SOUZA LTDA', 'SUPERMERCADO SOUZA', 'PRIMAVERA DO LESTE', 'MT', '03.353.940/0001-13', '66 34981152'),
    ('134', 'TRANSP. COM. E PREST. DE SERV. GUINCHO E MUNCK EIRELI', 'VALLE GUINCHOS E MUNCK', 'RONDONOPOLIS', 'MT', '18.824.288/0001-12', '66 34236611'),
    ('17', 'SUPER PESADOS LOCACOES E REMOCOES EIRELI', 'GUINCHOS E GUINDASTES SUPER PESADOS', 'RONDONOPOLIS - JACIARA', 'MT', '03.805.518/0001-92', '66 3422-2020'),
    ('149', 'R. G. - GUINCHOS E GUINDASTES EIRELI - ME', 'SOCORRO E GUINCHO SILVA', 'SINOP', 'MT', '15.228.604/0001-52', '66 35316499'),
    ('15', 'K GEOVANIO PEREIRA ME', 'LOGISTICA K GP', 'ALTAMIRA - ITANHAEM - ANAPU - MARABA', 'PA', '09.431.111/0001-56', '93 3515-2000'),
    ('177', 'SERVICOS E REMOCOES DE VEICULOS E MÁQUINAS - LTDA ME', 'GUINCHO E REMOCOES PARAGOMINAS', 'PARAGOMINAS', 'PA', '11.455.517/0001-38', '91 37293000'),
    ('87', 'G H MANUTENÇÃO E REPARAÇÃO DE MÁQUINAS E EQUI. LTDA', 'G H MANUTENÇÃO E REPARAÇÃO', 'CABEDELO - JOÃO PESSOA', 'PB', '31.258.948/0001-96', '83 3228-5000'),
    ('178', 'JOSE CARLOS DOS SANTOS 01222956401', 'JOSE CARLOS DOS SANTOS', 'CABEDELO - JOÃO PESSOA', 'PB', '11.831.393/0001-02', '83 32281200'),
    ('68', 'V M NOGUEIRA E CIA LTDA ME', 'AUTO SOCORRO MOURA', 'CARUARU - GRAVATA - BEZERROS', 'PE', '07.391.802/0001-91', '81 3721-4000'),
    ('185', 'RODOVIARIO CARUARUENSE LTDA', 'RODOVIARIO CARUARUENSE LTDA', 'CARUARU - PE', 'PE', '10.748.100/0001-12', '81 37277000'),
    ('3', 'LOCAGUINCHO LOCACOES E GUINCHOS LTDA', 'LOCAGUINCHO', 'RECIFE - CABO DE SANTO AGOSTINHO', 'PE', '03.023.238/0001-70', '81 3453-1000'),
    ('133', 'CONSTRUTORA E TRANSPORTE FREITAS LTDA', 'CONSTRUTORA FREITAS', 'SALGUEIRO', 'PE', '10.825.100/0001-98', '87 38712000'),
    ('13', 'G D D OLIVEIRA SERVICO DE GUINCHO', 'AUTO SOCORRO OLIVEIRA', 'PARANAGUA - MATINHOS - MORRETES - ANTONINA', 'PR', '08.810.150/0001-08', '41 3423-1122'),
    ('124', 'AGROPECUARIA E TRANSPORTES SCHLOSSER EIRELI', 'SCHLOSSER TRANSPORTES', 'CASCAVEL', 'PR', '32.188.700/0001-65', '45 32252000'),
    ('11', 'F M S GUINCHOS E TRANSPORTES EIRELI', 'F M S GUINCHOS', 'CURITIBA - PALMEIRA', 'PR', '07.288.510/0001-52', '41 3345-8000'),
    ('12', 'H L B - GUINCHOS E SOCORRO EIRELI', 'HLB GUINCHOS', 'CURITIBA - SAO JOSE DOS PINHAIS', 'PR', '08.120.300/0001-90', '41 3282-5000'),
    ('123', 'O L DA CUNHA GUINCHO', 'MUNCK E GUINCHO CUNHA', 'FOZ DO IGUACU', 'PR', '31.988.100/0001-20', '45 35251000'),
    ('142', 'C F C ROLANDIA SOCIEDADE SIMPLES UNIPESSOAL LTDA', 'CFC ROLANDIA', 'LONDRINA - ROLANDIA - ARAPONGAS', 'PR', '41.120.300/0001-88', '43 32561200'),
    ('105', 'N L S DE OLIVEIRA GUINCHOS', 'N L S GUINCHOS', 'MARINGA - SARANDI', 'PR', '33.820.100/0001-15', '44 32623000'),
    ('122', 'M R S TRANSPORTES E GUINCHOS EIRELI', 'M R S GUINCHOS', 'PONTA GROSSA', 'PR', '31.850.400/0001-30', '42 32244000'),
    ('16', 'L. C. D. DA SILVA GUINCHO', 'AUTO SOCORRO SILVA', 'CAMPOS DOS GOYTACAZES', 'RJ', '09.520.110/0001-40', '22 2733-1000'),
    ('8', 'CRANE & HEAVY LIFTING TRANSPORTES EIRELI', 'CRANE & HEAVY LIFTING', 'DUQUE DE CAXIAS - MACAE', 'RJ', '05.810.200/0001-60', '21 2671-5000'),
    ('10', 'B R S SERVICOS DE GUINCHO E TRANSPORTES EIRELI', 'BRS GUINCHOS', 'ITAGUAI - SEROPEDICA', 'RJ', '06.920.400/0001-80', '21 2688-3000'),
    ('125', 'J A DE OLIVEIRA REMOCOES E GUINCHOS', 'J A REMOCOES', 'MACAE - RIO DAS OSTRAS', 'RJ', '33.120.500/0001-70', '22 2772-1500'),
    ('7', 'E G M GUINCHOS E TRANSPORTES EIRELI', 'EGM GUINCHOS', 'RESENDE - PORTO REAL - ITATIAIA', 'RJ', '04.910.800/0001-30', '24 3354-2000'),
    ('6', 'A M S GUINCHOS E LOGISTICA EIRELI', 'AMS LOGISTICA E GUINCHOS', 'RIO DE JANEIRO - NITERÓI', 'RJ', '04.120.900/0001-10', '21 2580-4000'),
    ('9', 'G S T GUINCHOS E REMOCOES EIRELI', 'GST REMOCOES', 'VOLTA REDONDA - BARRA MANSA', 'RJ', '06.110.300/0001-50', '24 3348-1000'),
    ('126', 'M S DE SOUZA GUINCHOS E TRANSPORTES', 'M S SOUZA GUINCHOS', 'MOSSORO - ASSU', 'RN', '34.250.100/0001-40', '84 3316-2000'),
    ('183', 'N A TRANSPORTES E LOGISTICA EIRELI', 'N A LOGISTICA', 'NATAL - PARNAMIRIM', 'RN', '21.820.600/0001-90', '84 3206-5000'),
    ('127', 'F A DE LIMA GUINCHOS', 'F A LIMA GUINCHOS', 'PORTO VELHO', 'RO', '35.110.200/0001-20', '69 3225-4000'),
    ('182', 'J R S TRANSPORTES E GUINCHOS EIRELI', 'J R S GUINCHOS', 'BOA VISTA', 'RR', '20.120.400/0001-10', '95 3623-1000'),
    ('128', 'R S DE OLIVEIRA GUINCHOS', 'R S GUINCHOS', 'CAXIAS DO SUL - FARROUPILHA', 'RS', '36.120.800/0001-80', '54 3222-1000'),
    ('5', 'E R S GUINCHOS E TRANSPORTES EIRELI', 'ERS GUINCHOS', 'PASSO FUNDO - ERECHIM', 'RS', '03.920.700/0001-20', '54 3313-5000'),
    ('2', 'G M S LOGISTICA E GUINCHOS EIRELI', 'GMS LOGISTICA', 'PORTO ALEGRE - CANOAS', 'RS', '02.110.400/0001-90', '51 3342-2000'),
    ('4', 'H R T GUINCHOS E REMOCOES EIRELI', 'HRT REMOCOES', 'RIO GRANDE - PELOTAS', 'RS', '03.120.500/0001-10', '53 3231-4000'),
    ('1', 'A B C GUINCHOS E TRANSPORTES LTDA', 'ABC GUINCHOS', 'SANTA MARIA', 'RS', '01.234.567/0001-89', '55 3221-1234'),
    ('130', 'T R S GUINCHOS E TRANSPORTES', 'T R S GUINCHOS', 'CHAPECO - XANXERE', 'SC', '37.820.300/0001-50', '49 3322-8000'),
    ('181', 'K L M GUINCHOS E LOGISTICA EIRELI', 'K L M LOGISTICA', 'CRICIUMA - IÇARA', 'SC', '19.820.100/0001-30', '48 3433-9000'),
    ('140', 'L M S GUINCHOS E REMOCOES EIRELI', 'L M S REMOCOES', 'FLORIANOPOLIS - JOSE', 'SC', '25.120.400/0001-60', '48 3244-1000'),
    ('141', 'P Q R GUINCHOS E TRANSPORTES EIRELI', 'P Q R GUINCHOS', 'ITAJAI - NAVEGANTES', 'SC', '26.820.700/0001-90', '47 3348-3000'),
    ('137', 'V W X GUINCHOS E LOGISTICA EIRELI', 'V W X GUINCHOS', 'JOINVILLE', 'SC', '23.120.800/0001-20', '47 3433-2000'),
    ('136', 'Y Z A GUINCHOS E REMOCOES EIRELI', 'Y Z A REMOCOES', 'ARACAJU - SOCORRO', 'SE', '22.820.500/0001-40', '79 3214-5000'),
    ('119', 'A A A GUINCHOS E TRANSPORTES EIRELI', 'A A A GUINCHOS', 'AMERICANA - SANTA BARBARA', 'SP', '33.110.400/0001-10', '19 3461-8000'),
    ('116', 'B B B GUINCHOS E REMOCOES EIRELI', 'B B B REMOCOES', 'ARACATUBA - BIRIGUI', 'SP', '30.120.700/0001-30', '18 3623-4000'),
    ('114', 'C C C GUINCHOS E LOGISTICA EIRELI', 'C C C LOGISTICA', 'ARARAQUARA - SÃO CARLOS', 'SP', '28.820.200/0001-50', '16 3332-1000'),
    ('113', 'D D D GUINCHOS E TRANSPORTES EIRELI', 'D D D GUINCHOS', 'BAURU - JAÚ', 'SP', '27.120.900/0001-80', '14 3234-5000'),
    ('111', 'E E E GUINCHOS E REMOCOES EIRELI', 'E E E REMOCOES', 'CAMPINAS - SUMARÉ - HORTOLÂNDIA', 'SP', '26.820.300/0001-10', '19 3251-2000'),
    ('98', 'F F F GUINCHOS E LOGISTICA EIRELI', 'F F F GUINCHOS', 'FRANCA', 'SP', '18.120.500/0001-40', '16 3722-3000'),
    ('97', 'G G G GUINCHOS E TRANSPORTES EIRELI', 'G G G REMOCOES', 'GUARULHOS - ARUJA', 'SP', '17.820.800/0001-70', '11 2408-4000'),
    ('95', 'H H H GUINCHOS E LOGISTICA EIRELI', 'H H H LOGISTICA', 'ITAPEVICA - BARUERI', 'SP', '15.120.100/0001-90', '11 4191-5000'),
    ('93', 'I I I GUINCHOS E TRANSPORTES EIRELI', 'I I I GUINCHOS', 'JUNDIAI - ITUPEVA', 'SP', '13.820.400/0001-20', '11 4586-6000'),
    ('92', 'J J J GUINCHOS E REMOCOES EIRELI', 'J J J REMOCOES', 'LIMEIRA - CORDEIROPOLIS', 'SP', '12.120.700/0001-40', '19 3451-7000'),
    ('91', 'K K K GUINCHOS E LOGISTICA EIRELI', 'K K K LOGISTICA', 'MARILIA', 'SP', '11.820.900/0001-60', '14 3433-8000'),
    ('90', 'L L L GUINCHOS E TRANSPORTES EIRELI', 'L L L GUINCHOS', 'MOGI DAS CRUZES - SUZANO', 'SP', '10.120.200/0001-80', '11 4799-9000'),
    ('89', 'M M M GUINCHOS E REMOCOES EIRELI', 'M M M REMOCOES', 'PIRACICABA', 'SP', '09.820.500/0001-10', '19 3422-1000'),
    ('88', 'N N N GUINCHOS E LOGISTICA EIRELI', 'N N N LOGISTICA', 'PRESIDENTE PRUDENTE', 'SP', '08.120.800/0001-30', '18 3221-2000'),
    ('86', 'O O O GUINCHOS E TRANSPORTES EIRELI', 'O O O GUINCHOS', 'RIBEIRAO PRETO - SERTÃOZINHO', 'SP', '07.820.100/0001-50', '16 3636-3000'),
    ('82', 'P P P GUINCHOS E REMOCOES EIRELI', 'P P P REMOCOES', 'SANTO ANDRE - SÃO BERNARDO', 'SP', '06.120.400/0001-70', '11 4438-4000'),
    ('81', 'Q Q Q GUINCHOS E LOGISTICA EIRELI', 'Q Q Q LOGISTICA', 'SANTOS - CUBATÃO', 'SP', '05.820.700/0001-90', '13 3232-5000'),
    ('80', 'R R R GUINCHOS E TRANSPORTES EIRELI', 'R R R GUINCHOS', 'SAO JOSE DO RIO PRETO', 'SP', '04.120.900/0001-10', '17 3233-6000'),
    ('77', 'S S S GUINCHOS E REMOCOES EIRELI', 'S S S REMOCOES', 'SAO JOSE DOS CAMPOS - JACAREÍ', 'SP', '03.820.200/0001-30', '12 3921-7000'),
    ('76', 'T T T GUINCHOS E LOGISTICA EIRELI', 'T T T LOGISTICA', 'SAO PAULO (TODAS AS REGIOES)', 'SP', '02.120.500/0001-50', '11 3100-8000'),
    ('75', 'U U U GUINCHOS E TRANSPORTES EIRELI', 'U U U GUINCHOS', 'SOROCABA - VOTORANTIM', 'SP', '01.820.800/0001-70', '15 3231-9000'),
    ('74', 'V V V GUINCHOS E REMOCOES EIRELI', 'V V V REMOCOES', 'TAUBATE - PINDAMONHANGABA', 'SP', '00.120.100/0001-90', '12 3632-1000'),
    ('72', 'W W W GUINCHOS E LOGISTICA EIRELI', 'W W W LOGISTICA', 'VALINHOS - VINHEDO', 'SP', '34.820.300/0001-20', '19 3871-2000'),
    ('70', 'X X X GUINCHOS E TRANSPORTES EIRELI', 'X X X GUINCHOS', 'PALMAS - PORTO NACIONAL', 'TO', '32.120.600/0001-40', '63 3215-3000'),
    ('69', 'Y Y Y GUINCHOS E REMOCOES EIRELI', 'Y Y Y REMOCOES', 'ARAGUAINA', 'TO', '31.820.900/0001-60', '63 3412-4000'),
    ('66', 'Z Z Z GUINCHOS E LOGISTICA EIRELI', 'Z Z Z LOGISTICA', 'GURUPI', 'TO', '30.120.200/0001-80', '63 3312-5000'),
    ('65', 'A B1 GUINCHOS E TRANSPORTES EIRELI', 'A B1 GUINCHOS', 'MANAUS', 'AM', '29.820.500/0001-10', '92 3622-6000'),
    ('64', 'C D1 GUINCHOS E REMOCOES EIRELI', 'C D1 REMOCOES', 'MACAPA - SANTANA', 'AP', '28.120.800/0001-30', '96 3223-7000'),
    ('63', 'E F1 GUINCHOS E LOGISTICA EIRELI', 'E F1 LOGISTICA', 'RIO BRANCO', 'AC', '27.820.100/0001-50', '68 3224-8000'),
    ('62', 'G H1 GUINCHOS E TRANSPORTES EIRELI', 'G H1 GUINCHOS', 'MACEIO - RIO LARGO', 'AL', '26.120.400/0001-70', '82 3326-9000'),
    ('61', 'I J1 GUINCHOS E REMOCOES EIRELI', 'I J1 REMOCOES', 'ARACAJU', 'SE', '25.820.700/0001-90', '79 3215-1000'),
    ('60', 'K L1 GUINCHOS E LOGISTICA EIRELI', 'K L1 LOGISTICA', 'SAO LUIS - IMPERATRIZ', 'MA', '24.120.900/0001-10', '98 3235-2000'),
    ('59', 'M N1 GUINCHOS E TRANSPORTES EIRELI', 'M N1 GUINCHOS', 'TERESINA - PARNAIBA', 'PI', '23.820.200/0001-30', '86 3221-3000'),
    ('58', 'O P1 GUINCHOS E REMOCOES EIRELI', 'O P1 REMOCOES', 'CAMPO GRANDE - DOURADOS', 'MS', '22.120.500/0001-50', '67 3321-4000'),
    ('57', 'Q R1 GUINCHOS E LOGISTICA EIRELI', 'Q R1 LOGISTICA', 'CUIABA - VARZEA GRANDE', 'MT', '21.820.800/0001-70', '65 3623-5000'),
    ('56', 'S T1 GUINCHOS E TRANSPORTES EIRELI', 'S T1 GUINCHOS', 'GOIANIA - APARECIDA', 'GO', '20.120.100/0001-90', '62 3212-6000'),
    ('55', 'U V1 GUINCHOS E REMOCOES EIRELI', 'U V1 REMOCOES', 'BRASILIA - TAGUATINGA', 'DF', '19.820.400/0001-10', '61 3321-7000'),
    ('53', 'W X1 GUINCHOS E LOGISTICA EIRELI', 'W X1 LOGISTICA', 'BELO HORIZONTE - CONTAGEM', 'MG', '18.120.700/0001-30', '31 3212-8000'),
    ('52', 'Y Z1 GUINCHOS E TRANSPORTES EIRELI', 'Y Z1 GUINCHOS', 'VITORIA - VILA VELHA', 'ES', '17.820.000/0001-50', '27 3322-9000'),
    ('50', 'A B2 GUINCHOS E REMOCOES EIRELI', 'A B2 REMOCOES', 'RIO DE JANEIRO - NITEROI', 'RJ', '16.120.300/0001-70', '21 2212-1000'),
    ('47', 'C D2 GUINCHOS E LOGISTICA EIRELI', 'C D2 LOGISTICA', 'SAO PAULO - GUARULHOS', 'SP', '15.820.600/0001-90', '11 2112-2000'),
    ('46', 'E F2 GUINCHOS E TRANSPORTES EIRELI', 'E F2 GUINCHOS', 'CURITIBA - LONDRINA', 'PR', '14.120.900/0001-10', '41 3112-3000'),
    ('45', 'G H2 GUINCHOS E REMOCOES EIRELI', 'G H2 REMOCOES', 'FLORIANOPOLIS - JOINVILLE', 'SC', '13.820.200/0001-30', '48 3112-4000'),
    ('44', 'I J2 GUINCHOS E LOGISTICA EIRELI', 'I J2 LOGISTICA', 'PORTO ALEGRE - CAXIAS', 'RS', '12.120.500/0001-50', '51 3112-5000'),
    ('42', 'K L2 GUINCHOS E TRANSPORTES EIRELI', 'K L2 GUINCHOS', 'SALVADOR - FEIRA DE SANTANA', 'BA', '11.820.800/0001-70', '71 3112-6000'),
    ('41', 'M N2 GUINCHOS E REMOCOES EIRELI', 'M N2 REMOCOES', 'RECIFE - OLINDA', 'PE', '10.120.100/0001-90', '81 3112-7000'),
    ('40', 'O P2 GUINCHOS E LOGISTICA EIRELI', 'O P2 LOGISTICA', 'FORTALEZA - CAUCAIA', 'CE', '09.820.400/0001-10', '85 3112-8000'),
    ('39', 'Q R2 GUINCHOS E TRANSPORTES EIRELI', 'Q R2 GUINCHOS', 'BELEM - ANANINDEUA', 'PA', '08.120.700/0001-30', '91 3112-9000')
]

@st.cache_resource
def init_db():
    with get_connection() as conn:
        with conn.cursor() as cursor:
            # Tabela de Usuários
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    usuario TEXT UNIQUE NOT NULL,
                    senha TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    perfil TEXT NOT NULL
                )
            """)

            usuarios_iniciais = [
                ("admin", "Transpes@1966", "Administrador do Sistema", "ADMIN"),
                ("deyves.teixeira", "Transpes@26", "Deyves Teixeira", "ADMIN"),
                ("felipe.cesario", "Transpes@26", "Felipe Cesario", "ADMIN")
            ]

            for usr, pwd, nome, perf in usuarios_iniciais:
                cursor.execute("SELECT COUNT(*) FROM usuarios WHERE usuario = %s", (usr,))
                if cursor.fetchone()[0] == 0:
                    cursor.execute(
                        "INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (%s, %s, %s, %s)",
                        (usr, hash_senha(pwd), nome, perf)
                    )
                else:
                    cursor.execute(
                        "UPDATE usuarios SET senha = %s WHERE usuario = %s",
                        (hash_senha(pwd), usr)
                    )

            # Tabela de Cargas
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cargas (
                    id SERIAL PRIMARY KEY,
                    numero_carga TEXT UNIQUE NOT NULL,
                    cliente_origem TEXT DEFAULT '',
                    cliente_destino TEXT DEFAULT '',
                    nome_motorista TEXT NOT NULL,
                    cpf_motorista TEXT NOT NULL,
                    telefone_motorista TEXT,
                    tipo_veiculo TEXT,
                    placa_cavalo TEXT NOT NULL,
                    placa_carreta TEXT,
                    quantidade_eixos INTEGER NOT NULL,
                    peso_total REAL NOT NULL,
                    tipo_carga TEXT NOT NULL,
                    medida_dn TEXT,
                    valor_rpa REAL NOT NULL,
                    tipo_motorista TEXT NOT NULL,
                    cidade_origem TEXT NOT NULL,
                    estado_origem TEXT NOT NULL,
                    cidade_destino TEXT NOT NULL,
                    estado_destino TEXT NOT NULL,
                    data_carregamento TEXT NOT NULL,
                    previsao_descarga TEXT NOT NULL,
                    origens_json TEXT,
                    destinos_json TEXT,
                    tem_troca_nota INTEGER DEFAULT 0,
                    cidade_troca_nota TEXT,
                    estado_troca_nota TEXT,
                    data_troca_nota TEXT,
                    status TEXT DEFAULT 'PROGRAMADA',
                    data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Tabela Carga Expedição
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_expedicao (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    numero_set TEXT,
                    numero_viagem TEXT,
                    numero_cte TEXT,
                    numero_mdfe TEXT,
                    numero_nota_fiscal TEXT,
                    valor_pedagio_pago REAL DEFAULT 0.00,
                    data_saida_filial TEXT NOT NULL,
                    observacoes_expedicao TEXT,
                    data_emissao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Tabela Carga Operacional
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_operacional (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    data_descarga TEXT NOT NULL,
                    receita_frete REAL NOT NULL,
                    receita_pedagio REAL DEFAULT 0.00,
                    receita_taxa_descarga REAL DEFAULT 0.00,
                    fornecedor_descarga TEXT,
                    equipamento_descarga TEXT,
                    custo_fornecedor_descarga REAL DEFAULT 0.00,
                    peso_descarregado REAL,
                    observacoes_descarga TEXT,
                    receita_total REAL,
                    custo_total REAL,
                    margem_lucro_reais REAL,
                    margem_lucro_pct REAL,
                    data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Tabela Carga Administração
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_administracao (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    valor_adiantamento REAL,
                    valor_saldo REAL,
                    comprovante_entregue INTEGER DEFAULT 0,
                    data_liberacao_saldo TEXT,
                    status_pagamento_saldo TEXT DEFAULT 'PENDENTE'
                )
            """)

            # Nova Tabela: Fornecedores
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS fornecedores (
                    id SERIAL PRIMARY KEY,
                    nr_contrato TEXT,
                    razao_social TEXT NOT NULL,
                    nome_fantasia TEXT,
                    local_atendimento TEXT,
                    uf TEXT,
                    cpf_cnpj TEXT,
                    contato TEXT,
                    data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Carga Inicial dos Fornecedores da Planilha (se a tabela estiver vazia)
            cursor.execute("SELECT COUNT(*) FROM fornecedores")
            if cursor.fetchone()[0] == 0:
                for row in FORNECEDORES_INICIAIS:
                    cursor.execute("""
                        INSERT INTO fornecedores (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """, row)

            conn.commit()

init_db()

# ==========================================
# FUNÇÕES AUXILIARES DE EXPORTAÇÃO E IMAGEM
# ==========================================
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def gerar_numero_carga_novo():
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM cargas")
            qtd = cursor.fetchone()[0] + 1
    ano = datetime.datetime.now().year
    return f"TRP-{ano}-{qtd:03d}"

def gerar_excel_geral(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Visao_Geral_Cargas')
        worksheet = writer.sheets['Visao_Geral_Cargas']
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
    output.seek(0)
    return output

def gerar_pdf_geral(df):
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    elements = []
    styles = getSampleStyleSheet()
    
    titulo_style = ParagraphStyle(
        'TituloPDF',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#0E2F56'),
        alignment=1,
        spaceAfter=15
    )
    
    elements.append(Paragraph("Relatório Geral de Cargas - Transpes", titulo_style))
    elements.append(Spacer(1, 10))
    
    colunas_pdf = [
        "Nº Carga", "Status", "Motorista", "Placa Cavalo", 
        "Origem", "Destino", "Receita (R$)", "Custo (R$)", "Margem (R$)", "Margem (%)"
    ]
    
    cols_existentes = [c for c in colunas_pdf if c in df.columns]
    df_pdf = df[cols_existentes].copy()
    
    table_data = [cols_existentes]
    for row in df_pdf.itertuples(index=False):
        table_data.append([str(val) if val is not None else "" for val in row])
        
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0E2F56')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9FA')])
    ]))
    
    elements.append(t)
    doc.build(elements)
    output.seek(0)
    return output

# ==========================================
# DIÁLOGOS / POP-UPS DE CONFIRMAÇÃO E CADASTRO
# ==========================================
@st.dialog("📌 Carga Programada com Sucesso!")
def exibir_popup_programacao(num_carga):
    st.success(f"A carga **{num_carga}** foi cadastrada e seu status está como **PROGRAMADA**!")
    st.write("Os dados foram salvos no sistema e o formulário foi limpo para a próxima programação.")
    if st.button("OK / Próxima Programação", type="primary", use_container_width=True):
        st.session_state["exibir_modal_programacao"] = False
        st.session_state["prog_form_version"] += 1
        st.rerun()

@st.dialog("📦 Carga Expedida com Sucesso!")
def exibir_popup_expedicao(num_carga):
    st.success(f"A carga **{num_carga}** foi expedida e atualizada para **EM TRÂNSITO**!")
    st.write("Os dados foram salvos no sistema e os campos foram limpos para o próximo processo.")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_expedicao"] = False
        st.session_state["exp_form_version"] += 1
        st.rerun()

@st.dialog("⚙️ Carga Agendada para Descarga!")
def exibir_popup_operacional(num_carga):
    st.success(f"A carga **{num_carga}** teve o descarregamento registrado e seu status foi atualizado para **ENTREGUE**!")
    st.write("Informações financeiras operacionais gravadas com sucesso.")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_operacional"] = False
        st.session_state["op_form_version"] += 1
        st.rerun()

@st.dialog("💼 Processo Finalizado - Saldo Liberado com Sucesso!")
def exibir_popup_administracao(num_carga):
    st.success(f"A carga **{num_carga}** teve o acerto concluído e o status alterado para **FINALIZADA**!")
    st.write("A liberação de saldo foi registrada com sucesso.")
    if st.button("OK / Concluir", type="primary", use_container_width=True):
        st.session_state["exibir_modal_administracao"] = False
        st.session_state["adm_form_version"] += 1
        st.rerun()

@st.dialog("➕ Cadastrar Novo Fornecedor")
def modal_cadastrar_fornecedor():
    with st.form("form_modal_fornecedor"):
        f1, f2 = st.columns(2)
        nr_contrato = f1.text_input("Nr. Contrato").strip()
        cpf_cnpj = f2.text_input("CPF/CNPJ*").strip()

        razao_social = st.text_input("Razão Social*").strip().upper()
        nome_fantasia = st.text_input("Nome Fantasia").strip().upper()

        l1, l2 = st.columns([3, 1])
        local_atendimento = l1.text_input("Local de Atendimento*").strip().upper()
        uf = l2.text_input("UF*").strip().upper()

        contato = st.text_input("Contato / Telefone").strip()

        st.markdown("<br>", unsafe_allow_html=True)
        btn_salvar_forn = st.form_submit_button("💾 Salvar Fornecedor", type="primary", use_container_width=True)

    if btn_salvar_forn:
        if not razao_social or not local_atendimento or not uf:
            st.error("Por favor, preencha os campos obrigatórios (Razão Social, Local de Atendimento e UF).")
        else:
            try:
                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute("""
                            INSERT INTO fornecedores (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """, (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato))
                        conn.commit()
                st.success("Fornecedor cadastrado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao cadastrar fornecedor: {e}")

# ==========================================
# AUTENTICAÇÃO E SESSÃO (TELA DE LOGIN)
# ==========================================
if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    try:
        bg_base64 = get_base64_image("fundo_transpes.png")
        bg_style = f"""
            background-image: url("data:image/jpeg;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        """
    except Exception:
        bg_style = "background-color: #0E2F56;"

    try:
        logo_base64 = get_base64_image("logo_transpes.png")
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-width: 180px; margin-bottom: 10px;">'
    except Exception:
        logo_html = '<div class="login-title" style="color: #0E2F56; font-size: 26px; font-weight: bold;">ERP Transpes</div>'

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"], [data-testid="stHeader"] {{
            display: none;
        }}
        .stApp {{
            {bg_style}
        }}
        [data-testid="stMainBlockContainer"] {{
            max-width: 420px !important;
            padding-top: 5rem !important;
            padding-bottom: 2rem !important;
            margin: auto !important;
        }}
        .login-card {{
            background: rgba(255, 255, 255, 0.95);
            padding: 25px 20px 18px 20px;
            border-radius: 14px;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
            text-align: center;
            margin-bottom: 20px;
        }}
        .login-subtitle {{
            color: #444444;
            font-size: 13px;
            font-weight: 500;
        }}
        [data-testid="stWidgetLabel"] p {{
            color: #FFFFFF !important;
            font-weight: bold !important;
            font-size: 15px !important;
            text-shadow: 1px 1px 3px rgba(0,0,0,0.8);
        }}
        [data-testid="stNotification"] {{
            background-color: rgba(255, 235, 235, 0.95) !important;
            border-left: 5px solid #FF0000 !important;
        }}
        [data-testid="stNotification"] p {{
            color: #D32F2F !important;
            font-weight: bold !important;
            font-size: 15px !important;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="login-card">
            {logo_html}
            <div class="login-subtitle">Entre com suas credenciais de acesso</div>
        </div>
    """, unsafe_allow_html=True)

    with st.form("form_login"):
        usuario_input = st.text_input("Usuário", placeholder="Digite seu usuário")
        senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha")
        st.markdown("<br>", unsafe_allow_html=True)
        btn_login = st.form_submit_button("Entrar", use_container_width=True, type="primary")

    if btn_login:
        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT usuario, nome, perfil FROM usuarios WHERE usuario = %s AND senha = %s",
                    (usuario_input.strip().lower(), hash_senha(senha_input))
                )
                usr = cursor.fetchone()
        
        if usr:
            st.session_state["logado"] = True
            st.session_state["usuario"] = usr[0]
            st.session_state["nome"] = usr[1]
            st.session_state["perfil"] = usr[2]
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")

    st.stop()

# ==========================================
# BARRA LATERAL E NAVEGAÇÃO
# ==========================================
st.sidebar.markdown(f"### 👤 {st.session_state['nome']}")
st.sidebar.caption(f"Perfil: **{st.session_state['perfil']}**")

if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.clear()
    st.rerun()

perfil = st.session_state["perfil"]

opcoes = ["Visão Geral"]
icones = ["bar-chart"]

if perfil == "ADMIN":
    opcoes.extend(["Programação", "Expedição", "Operacional", "Administração", "Fornecedores", "Excluir Cargas", "Usuários"])
    icones.extend(["clipboard-plus", "file-earmark-text", "tools", "briefcase", "truck", "trash", "people"])
else:
    if perfil == "PROGRAMACAO":
        opcoes.append("Programação")
        icones.append("clipboard-plus")
    elif perfil == "EXPEDICAO":
        opcoes.append("Expedição")
        icones.append("file-earmark-text")
    elif perfil == "OPERACIONAL":
        opcoes.extend(["Operacional", "Fornecedores"])
        icones.extend(["tools", "truck"])
    elif perfil == "ADMINISTRATIVO":
        opcoes.extend(["Administração", "Fornecedores"])
        icones.extend(["briefcase", "truck"])

with st.sidebar:
    menu_selecionado = option_menu(
        "Menu Principal",
        opcoes,
        icons=icones,
        menu_icon="cast",
        default_index=0
    )

# ==========================================
# PÁGINAS DO SISTEMA
# ==========================================

# 1. VISÃO GERAL
if menu_selecionado == "Visão Geral":
    st.title("📊 Visão Geral e Relatórios")
    
    query = """
        SELECT 
            c.numero_carga AS "Nº Carga",
            c.status AS "Status",
            c.nome_motorista AS "Motorista",
            c.placa_cavalo AS "Placa Cavalo",
            c.cidade_origem || '/' || c.estado_origem AS "Origem",
            c.cidade_destino || '/' || c.estado_destino AS "Destino",
            
            COALESCE(o.receita_frete, 0.0) AS "Receita Frete (R$)",
            COALESCE(o.receita_pedagio, 0.0) AS "Receita Pedágio (R$)",
            COALESCE(o.receita_taxa_descarga, 0.0) AS "Receita Taxa Descarga (R$)",
            (COALESCE(o.receita_frete, 0.0) + COALESCE(o.receita_pedagio, 0.0) + COALESCE(o.receita_taxa_descarga, 0.0)) AS "Receita Total (R$)",
            
            COALESCE(c.valor_rpa, 0.0) AS "RPA (R$)",
            COALESCE(e.valor_pedagio_pago, 0.0) AS "Pedágio Pago (R$)",
            COALESCE(o.custo_fornecedor_descarga, 0.0) AS "Custo Descarga (R$)",
            (COALESCE(c.valor_rpa, 0.0) + COALESCE(e.valor_pedagio_pago, 0.0) + COALESCE(o.custo_fornecedor_descarga, 0.0)) AS "Custo Total (R$)"
            
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        LEFT JOIN carga_operacional o ON c.id = o.carga_id
        LEFT JOIN carga_administracao a ON c.id = a.carga_id
        ORDER BY c.id DESC
    """
    with get_connection() as conn:
        df = pd.read_sql_query(query, conn)

    if df.empty:
        st.info("Nenhuma carga cadastrada até o momento.")
    else:
        def formatar_real(valor):
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        df["Margem (R$)"] = df["Receita Total (R$)"] - df["Custo Total (R$)"]
        df["Margem (%)"] = df.apply(
            lambda r: (r["Margem (R$)"] / r["Receita Total (R$)"] * 100) if r["Receita Total (R$)"] > 0 else 0.0, 
            axis=1
        )

        rec_tot = df["Receita Total (R$)"].sum()
        custo_tot = df["Custo Total (R$)"].sum()
        margem_tot = df["Margem (R$)"].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Cargas", len(df))
        m2.metric("Receita Total", formatar_real(rec_tot))
        m3.metric("Custos Totais", formatar_real(custo_tot))
        m4.metric("Margem Total", formatar_real(margem_tot))
        m5.metric("Margem Média", f"{(margem_tot / rec_tot * 100) if rec_tot > 0 else 0.0:.1f}%")

        st.markdown("##### ⏳ Pendências por Setor")
        
        pend_expedicao = len(df[df["Status"] == "PROGRAMADA"])
        pend_operacional = len(df[df["Status"] == "EM TRÂNSITO"])
        pend_administracao = len(df[df["Status"] == "ENTREGUE"])

        p1, p2, p3 = st.columns(3)
        p1.metric("📦 Pendências EXPEDIÇÃO", f"{pend_expedicao} carga(s)", delta="Aguardando CT-e / Saída", delta_color="off")
        p2.metric("⚙️ Pendências OPERACIONAL", f"{pend_operacional} carga(s)", delta="Em trânsito / A descarregar", delta_color="off")
        p3.metric("💼 Pendências ADMINISTRATIVO", f"{pend_administracao} carga(s)", delta="Aguardando acerto financeiro", delta_color="off")

        st.markdown("---")
        
        st.subheader("📥 Exportar Relatório")
        col_exp1, col_exp2, _ = st.columns([1, 1, 2])

        with col_exp1:
            excel_data = gerar_excel_geral(df)
            st.download_button(
                label="📗 Exportar para Excel (.xlsx)",
                data=excel_data,
                file_name="Relatorio_Geral_Transpes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_exp2:
            pdf_data = gerar_pdf_geral(df)
            st.download_button(
                label="📕 Exportar para PDF",
                data=pdf_data,
                file_name="Relatorio_Geral_Transpes.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        st.markdown("---")
        st.subheader("📋 Detalhamento das Cargas")
        
        df_exibicao = df.copy()
        colunas_financeiras = [
            "Receita Total (R$)", "RPA (R$)", "Pedágio Pago (R$)", 
            "Custo Descarga (R$)", "Custo Total (R$)", "Margem (R$)"
        ]
        
        for col in colunas_financeiras:
            df_exibicao[col] = df_exibicao[col].apply(formatar_real)

        df_exibicao["Margem (%)"] = df_exibicao["Margem (%)"].apply(lambda v: f"{v:.1f}%")

        colunas_ordenadas = [
            "Nº Carga", "Status", "Motorista", "Origem", "Destino",
            "Receita Total (R$)", "RPA (R$)", "Pedágio Pago (R$)", 
            "Custo Descarga (R$)", "Custo Total (R$)", "Margem (R$)", "Margem (%)"
        ]
        
        st.dataframe(df_exibicao[colunas_ordenadas], use_container_width=True)

# 2. PROGRAMAÇÃO
elif menu_selecionado == "Programação":
    st.title("📌 Programação de Cargas")

    if "prog_form_version" not in st.session_state:
        st.session_state["prog_form_version"] = 0

    v = st.session_state["prog_form_version"]
    
    if st.session_state.get("exibir_modal_programacao", False):
        exibir_popup_programacao(st.session_state.get("carga_programada_num", ""))
    
    num_carga_sugerido = gerar_numero_carga_novo()

    col_qtd_orig, col_qtd_dest = st.columns(2)
    qtd_origens = col_qtd_orig.number_input("Qtd. de Clientes / Locais de Origem*", min_value=1, max_value=10, value=1, key=f"prog_qtd_origens_{v}")
    qtd_destinos = col_qtd_dest.number_input("Qtd. de Clientes / Locais de Destino*", min_value=1, max_value=10, value=1, key=f"prog_qtd_destinos_{v}")
    tem_troca_nota = st.checkbox("Houve Troca de Nota?", key=f"chk_troca_nota_{v}")

    with st.form(f"form_programacao_{v}"):
        st.subheader("1. Identificação da Carga")
        c1, _, _ = st.columns([1, 1, 1])
        numero_carga = c1.text_input("Número da Carga*", value=num_carga_sugerido).upper()

        st.subheader("2. Dados do Motorista e Veículo")
        m1, m2, m3 = st.columns(3)
        nome_motorista = m1.text_input("Nome do Motorista*").upper()
        cpf_motorista = m2.text_input("CPF do Motorista*")
        telefone_motorista = m3.text_input("Telefone")

        v1, v2, v3, v4, v5 = st.columns(5)
        tipo_motorista = v1.selectbox("Tipo de Motorista*", ["TERCEIRO", "FROTA", "AGREGADO"])
        tipo_veiculo = v2.selectbox("Tipo de Veículo*", ["CARRETA", "BITREM", "RODOTREM", "VANDERLÉIA", "TOCO", "TRUCK"])
        placa_cavalo = v3.text_input("Placa Cavalo*").upper()
        placa_carreta = v4.text_input("Placa Carreta").upper()
        qtd_eixos = v5.number_input("Qtd. Eixos*", min_value=2, max_value=9, value=6)

        st.subheader("3. Especificações da Carga")
        e1, e2, e3, e4 = st.columns(4)
        peso_total = e1.number_input("Peso Total (ton)*", min_value=0.1, value=30.0, step=0.5)
        tipo_carga = e2.text_input("Tipo da Carga*", value="GERAL").upper()
        medida_dn = e3.text_input("Medida DN").upper()
        valor_rpa = e4.number_input("Valor RPA (R$)*", min_value=0.0, value=0.0, step=100.0)

        st.subheader("4. Rotas e Clientes")
        
        lista_origens = []
        st.markdown("**📍 Locais e Clientes de Origem:**")
        for i in range(int(qtd_origens)):
            co1, co2, co3 = st.columns([2, 2, 1])
            cli_orig = co1.text_input(f"Cliente Origem {i+1}*", key=f"cli_orig_{i}_{v}").upper()
            cid_orig = co2.text_input(f"Cidade Origem {i+1}*", key=f"cid_orig_{i}_{v}").upper()
            est_orig = co3.text_input(f"UF Origem {i+1}*", key=f"est_orig_{i}_{v}").upper()
            lista_origens.append({"cliente": cli_orig, "cidade": cid_orig, "estado": est_orig})

        lista_destinos = []
        st.markdown("**🏁 Locais e Clientes de Destino:**")
        for j in range(int(qtd_destinos)):
            cd1, cd2, cd3 = st.columns([2, 2, 1])
            cli_dest = cd1.text_input(f"Cliente Destino {j+1}*", key=f"cli_dest_{j}_{v}").upper()
            cid_dest = cd2.text_input(f"Cidade Destino {j+1}*", key=f"cid_dest_{j}_{v}").upper()
            est_dest = cd3.text_input(f"UF Destino {j+1}*", key=f"est_dest_{j}_{v}").upper()
            lista_destinos.append({"cliente": cli_dest, "cidade": cid_dest, "estado": est_dest})

        st.subheader("5. Datas e Troca de Nota")
        d1, d2 = st.columns(2)
        data_carregamento = d1.date_input("Data Carregamento*", datetime.date.today(), format="DD/MM/YYYY")
        previsao_descarga = d2.date_input("Previsão Descarga*", datetime.date.today() + datetime.timedelta(days=2), format="DD/MM/YYYY")

        cidade_troca = ""
        estado_troca = ""
        data_troca_str = ""

        if tem_troca_nota:
            st.markdown("**📋 Dados da Troca de Nota:**")
            tn1, tn2, tn3 = st.columns(3)
            cidade_troca = tn1.text_input("Cidade da Troca de Nota*").upper()
            estado_troca = tn2.text_input("Estado da Troca de Nota*").upper()
            data_troca = tn3.date_input("Data da Troca de Nota*", datetime.date.today(), format="DD/MM/YYYY")
            data_troca_str = data_troca.strftime("%d/%m/%Y")

        st.markdown("<br>", unsafe_allow_html=True)
        salvar = st.form_submit_button("💾 Salvar Programação", type="primary", use_container_width=True)

    if salvar:
        if not lista_origens or not lista_destinos:
            st.error("Adicione pelo menos uma origem e um destino.")
        else:
            try:
                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        primeira_origem = lista_origens[0]
                        primeiro_destino = lista_destinos[0]

                        cursor.execute("""
                            INSERT INTO cargas (
                                numero_carga, cliente_origem, cliente_destino, nome_motorista, cpf_motorista,
                                telefone_motorista, tipo_veiculo, placa_cavalo, placa_carreta, quantidade_eixos,
                                peso_total, tipo_carga, medida_dn, valor_rpa, tipo_motorista,
                                cidade_origem, estado_origem, cidade_destino, estado_destino,
                                data_carregamento, previsao_descarga, origens_json, destinos_json,
                                tem_troca_nota, cidade_troca_nota, estado_troca_nota, data_troca_nota, status
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'PROGRAMADA')
                        """, (
                            numero_carga, primeira_origem["cliente"], primeiro_destino["cliente"], 
                            nome_motorista, cpf_motorista, telefone_motorista, tipo_veiculo, 
                            placa_cavalo, placa_carreta, qtd_eixos, peso_total, tipo_carga, 
                            medida_dn, valor_rpa, tipo_motorista, primeira_origem["cidade"], 
                            primeira_origem["estado"], primeiro_destino["cidade"], primeiro_destino["estado"],
                            data_carregamento.strftime("%d/%m/%Y"), previsao_descarga.strftime("%d/%m/%Y"),
                            json.dumps(lista_origens), json.dumps(lista_destinos),
                            1 if tem_troca_nota else 0, cidade_troca, estado_troca, data_troca_str
                        ))
                        conn.commit()

                st.session_state["exibir_modal_programacao"] = True
                st.session_state["carga_programada_num"] = numero_carga
                st.rerun()

            except psycopg2.IntegrityError:
                st.error(f"Erro: O número de carga '{numero_carga}' já existe.")
            except Exception as e:
                st.error(f"Erro ao salvar no banco de dados: {e}")

# 3. EXPEDIÇÃO
elif menu_selecionado == "Expedição":
    st.title("📦 Expedição e Emissão de Documentos")
    
    if "exp_form_version" not in st.session_state:
        st.session_state["exp_form_version"] = 0

    v_exp = st.session_state["exp_form_version"]
    
    if st.session_state.get("exibir_modal_expedicao", False):
        exibir_popup_expedicao(st.session_state.get("carga_expedicao_num", ""))

    query_exp = """
        SELECT 
            c.id, c.numero_carga, c.nome_motorista, c.cliente_origem,
            c.cidade_origem, c.estado_origem, c.cliente_destino,
            c.cidade_destino, c.estado_destino, COALESCE(c.valor_rpa, 0.0) AS valor_rpa
        FROM cargas c
        WHERE c.status = 'PROGRAMADA'
    """
    with get_connection() as conn:
        cargas_df = pd.read_sql_query(query_exp, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga aguardando expedição no momento.")
    else:
        opcoes_cargas = {}
        dados_rpa = {}
        for _, row in cargas_df.iterrows():
            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_rpa[row['id']] = row['valor_rpa']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"exp_sel_carga_{v_exp}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        
        v_rpa = dados_rpa[carga_id]
        v_adiantamento_calc = v_rpa * 0.70

        e_c1, e_c2, e_c3, e_c4 = st.columns(4)
        qtd_viagens = e_c1.number_input("Qtd. Viagens", 1, 10, 1, key=f"exp_q_v_{v_exp}")
        qtd_nfs = e_c2.number_input("Qtd. Notas Fiscais", 1, 10, 1, key=f"exp_q_nf_{v_exp}")
        qtd_ctes = e_c3.number_input("Qtd. CT-es", 1, 10, 1, key=f"exp_q_cte_{v_exp}")
        qtd_mdfes = e_c4.number_input("Qtd. MDF-es", 1, 10, 1, key=f"exp_q_mdfe_{v_exp}")

        with st.form(f"form_expedicao_{v_exp}"):
            st.text_input("Valor Adiantamento (70% do RPA)", value=f"R$ {v_adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)

            numero_set = st.text_input("Número do SET*", placeholder="Ex: 55421").upper()

            st.subheader("🚩 Viagens")
            lista_viagens = []
            cols_v = st.columns(min(qtd_viagens, 4))
            for i in range(qtd_viagens):
                val_v = cols_v[i % 4].text_input(f"Viagem {i+1}*", key=f"exp_viagem_{i}_{v_exp}").upper()
                if val_v.strip():
                    lista_viagens.append(val_v.strip())

            st.subheader("📄 Notas Fiscais (NF)")
            lista_nfs = []
            cols_nf = st.columns(min(qtd_nfs, 4))
            for i in range(qtd_nfs):
                val_nf = cols_nf[i % 4].text_input(f"Nº Nota Fiscal {i+1}*", key=f"exp_nf_{i}_{v_exp}").upper()
                if val_nf.strip():
                    lista_nfs.append(val_nf.strip())

            st.subheader("📑 Conhecimentos de Transporte (CT-e)")
            lista_ctes = []
            cols_cte = st.columns(min(qtd_ctes, 4))
            for i in range(qtd_ctes):
                val_cte = cols_cte[i % 4].text_input(f"Nº CT-e {i+1}*", key=f"exp_cte_{i}_{v_exp}").upper()
                if val_cte.strip():
                    lista_ctes.append(val_cte.strip())

            st.subheader("📋 Manifestos (MDF-e)")
            lista_mdfes = []
            cols_mdfe = st.columns(min(qtd_mdfes, 4))
            for i in range(qtd_mdfes):
                val_mdfe = cols_mdfe[i % 4].text_input(f"Nº MDF-e {i+1}*", key=f"exp_mdfe_{i}_{v_exp}").upper()
                if val_mdfe.strip():
                    lista_mdfes.append(val_mdfe.strip())

            col_outros1, col_outros2 = st.columns(2)
            valor_pedagio_pago = col_outros1.number_input("Valor Pedágio Pago ao Motorista (R$)", min_value=0.0, value=0.0)
            data_saida = col_outros2.date_input("Data de Saída / Início Viagem*", datetime.date.today(), format="DD/MM/YYYY")

            obs_exp = st.text_area("Observações da Expedição").upper()

            salvar_exp = st.form_submit_button("🚚 Confirmar Saída / Enviar para Trânsito", use_container_width=True, type="primary")

        if salvar_exp:
            if not numero_set.strip():
                st.error("Por favor, preencha o Número do SET.")
            elif len(lista_viagens) < qtd_viagens:
                st.error("Por favor, preencha todos os campos de Viagem selecionados.")
            elif len(lista_nfs) < qtd_nfs:
                st.error("Por favor, preencha todos os campos de Nota Fiscal selecionados.")
            elif len(lista_ctes) < qtd_ctes:
                st.error("Por favor, preencha todos os campos de CT-e selecionados.")
            elif len(lista_mdfes) < qtd_mdfes:
                st.error("Por favor, preencha todos os campos de MDF-e selecionados.")
            else:
                json_viagens = json.dumps(lista_viagens)
                json_nfs = json.dumps(lista_nfs)
                json_ctes = json.dumps(lista_ctes)
                json_mdfes = json.dumps(lista_mdfes)

                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute("""
                            INSERT INTO carga_expedicao (
                                carga_id, numero_set, numero_viagem, numero_cte, numero_mdfe, numero_nota_fiscal,
                                valor_pedagio_pago, data_saida_filial, observacoes_expedicao
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (
                            carga_id, 
                            numero_set.strip(), 
                            json_viagens,
                            json_ctes, 
                            json_mdfes, 
                            json_nfs, 
                            valor_pedagio_pago, 
                            data_saida.strftime("%d/%m/%Y"), 
                            obs_exp
                        ))
                        
                        cursor.execute("UPDATE cargas SET status = 'EM TRÂNSITO' WHERE id = %s", (carga_id,))
                        conn.commit()

                st.session_state["exibir_modal_expedicao"] = True
                st.session_state["carga_expedicao_num"] = num_carga_sel
                st.rerun()

# 4. OPERACIONAL
elif menu_selecionado == "Operacional":
    st.title("⚙️ Operacional e Descarga")
    
    if "op_form_version" not in st.session_state:
        st.session_state["op_form_version"] = 0

    v_op = st.session_state["op_form_version"]

    if st.session_state.get("exibir_modal_operacional", False):
        exibir_popup_operacional(st.session_state.get("carga_operacional_num", ""))

    query_op = """
        SELECT 
            c.id, c.numero_carga, c.nome_motorista, c.cliente_origem,
            c.cidade_origem, c.estado_origem, c.cliente_destino,
            c.cidade_destino, c.estado_destino, c.previsao_descarga, e.numero_cte
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'EM TRÂNSITO'
    """
    with get_connection() as conn:
        cargas_df = pd.read_sql_query(query_op, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga em trânsito no momento.")
    else:
        opcoes_cargas = {}
        dados_cargas = {}
        for _, row in cargas_df.iterrows():
            cte_str = "N/A"
            if row['numero_cte']:
                try:
                    ctes = json.loads(row['numero_cte'])
                    cte_str = ", ".join(ctes) if isinstance(ctes, list) else str(ctes)
                except Exception:
                    cte_str = str(row['numero_cte'])

            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | CTe: {cte_str} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_cargas[row['id']] = row['previsao_descarga']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"op_sel_carga_{v_op}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        prev_descarga_val = dados_cargas[carga_id]

        # Carregar Lista de Fornecedores do Banco para o Autocomplete
        with get_connection() as conn:
            fornecedores_df = pd.read_sql_query("SELECT razao_social, nome_fantasia FROM fornecedores ORDER BY razao_social", conn)
        
        lista_opcoes_fornecedores = []
        for _, row in fornecedores_df.iterrows():
            nome = row['razao_social']
            if row['nome_fantasia']:
                nome += f" ({row['nome_fantasia']})"
            lista_opcoes_fornecedores.append(nome)

        with st.form(f"form_operacional_{v_op}"):
            st.markdown("---")
            o1, o2, o3 = st.columns(3)
            o1.text_input("Previsão de Descarga", value=prev_descarga_val, disabled=True)
            receita_frete = o2.number_input("Receita Frete (R$)*", min_value=0.0, value=0.0)
            receita_pedagio = o3.number_input("Receita Pedágio (R$)", min_value=0.0, value=0.0)

            o4, o5, o6 = st.columns(3)
            receita_taxa = o4.number_input("Receita Taxa Descarga (R$)", min_value=0.0, value=0.0)
            
            # Autocomplete de Fornecedores cadastrados
            if lista_opcoes_fornecedores:
                fornecedor = o5.selectbox("Fornecedor Descarga", [""] + lista_opcoes_fornecedores)
            else:
                fornecedor = o5.text_input("Fornecedor Descarga").upper()
                
            equipamento = o6.text_input("Equipamento Descarga").upper()

            custo_fornecedor = st.number_input("Custo Fornecedor (R$)", min_value=0.0, value=0.0)
            peso_descarregado = st.number_input("Peso Descarregado (ton)", min_value=0.0, value=0.0)
            obs = st.text_area("Observações").upper()

            salvar_op = st.form_submit_button("✅ Finalizar Operacional", use_container_width=True, type="primary")

        if salvar_op:
            receita_total = receita_frete + receita_pedagio + receita_taxa
            custo_total = custo_fornecedor
            margem_reais = receita_total - custo_total
            margem_pct = (margem_reais / receita_total * 100) if receita_total > 0 else 0.0

            with get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("""
                        INSERT INTO carga_operacional (
                            carga_id, data_descarga, receita_frete, receita_pedagio, receita_taxa_descarga,
                            fornecedor_descarga, equipamento_descarga, custo_fornecedor_descarga,
                            peso_descarregado, observacoes_descarga, receita_total, custo_total,
                            margem_lucro_reais, margem_lucro_pct
                        ) VALUES (%s, '', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (
                        carga_id, receita_frete, receita_pedagio, receita_taxa,
                        fornecedor, equipamento, custo_fornecedor, peso_descarregado, obs,
                        receita_total, custo_total, margem_reais, margem_pct
                    ))
                    
                    cursor.execute("UPDATE cargas SET status = 'ENTREGUE' WHERE id = %s", (carga_id,))
                    conn.commit()

            st.session_state["exibir_modal_operacional"] = True
            st.session_state["carga_operacional_num"] = num_carga_sel
            st.rerun()

# 5. ADMINISTRAÇÃO
elif menu_selecionado == "Administração":
    st.title("💼 Administração e Acerto Financeiro")
    
    if "adm_form_version" not in st.session_state:
        st.session_state["adm_form_version"] = 0

    v_adm = st.session_state["adm_form_version"]

    if st.session_state.get("exibir_modal_administracao", False):
        exibir_popup_administracao(st.session_state.get("carga_admin_num", ""))

    query_adm = """
        SELECT 
            c.id, c.numero_carga, c.nome_motorista, c.cliente_origem,
            c.cidade_origem, c.estado_origem, c.cliente_destino,
            c.cidade_destino, c.estado_destino, COALESCE(c.valor_rpa, 0.0) AS valor_rpa, e.numero_cte
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'ENTREGUE'
    """
    with get_connection() as conn:
        cargas_df = pd.read_sql_query(query_adm, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga entregue aguardando acerto administrativo.")
    else:
        opcoes_cargas = {}
        dados_rpa = {}
        for _, row in cargas_df.iterrows():
            cte_str = "N/A"
            if row['numero_cte']:
                try:
                    ctes = json.loads(row['numero_cte'])
                    cte_str = ", ".join(ctes) if isinstance(ctes, list) else str(ctes)
                except Exception:
                    cte_str = str(row['numero_cte'])

            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | CTe: {cte_str} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_rpa[row['id']] = row['valor_rpa']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"adm_sel_carga_{v_adm}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        
        v_rpa = dados_rpa[carga_id]
        adiantamento_calc = v_rpa * 0.70
        saldo_calc = v_rpa * 0.30

        with st.form(f"form_administracao_{v_adm}"):
            data_descarga_real = st.date_input("Data Efetiva da Descarga*", datetime.date.today(), format="DD/MM/YYYY")

            a1, a2 = st.columns(2)
            a1.text_input("Valor Adiantamento (70% RPA)", value=f"R$ {adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)
            a2.text_input("Valor Saldo (30% RPA)", value=f"R$ {saldo_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)

            a3, a4 = st.columns(2)
            comprovante = a3.checkbox("Comprovante de Entregue Recebido")
            status_pag = a4.selectbox("Status Pagamento Saldo", ["PENDENTE", "PAGO", "CANCELADO"])

            data_lib = st.date_input("Data Liberação Saldo", datetime.date.today(), format="DD/MM/YYYY")

            salvar_adm = st.form_submit_button("💰 Finalizar Acerto e Liberar Saldo", use_container_width=True, type="primary")

        if salvar_adm:
            with get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("""
                        UPDATE carga_operacional 
                        SET data_descarga = %s 
                        WHERE carga_id = %s
                    """, (data_descarga_real.strftime("%d/%m/%Y"), carga_id))

                    cursor.execute("""
                        INSERT INTO carga_administracao (
                            carga_id, valor_adiantamento, valor_saldo, comprovante_entregue,
                            data_liberacao_saldo, status_pagamento_saldo
                        ) VALUES (%s, %s, %s, %s, %s, %s)
                    """, (carga_id, adiantamento_calc, saldo_calc, 1 if comprovante else 0, data_lib.strftime("%d/%m/%Y"), status_pag))
                    
                    cursor.execute("UPDATE cargas SET status = 'FINALIZADA' WHERE id = %s", (carga_id,))
                    conn.commit()

            st.session_state["exibir_modal_administracao"] = True
            st.session_state["carga_admin_num"] = num_carga_sel
            st.rerun()

# 6. FORNECEDORES (NOVA ABA)
elif menu_selecionado == "Fornecedores":
    st.title("🚚 Gestão de Fornecedores")

    # Topo com Botão de Cadastro
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("➕ Cadastrar Novo Fornecedor", type="primary", use_container_width=True):
            modal_cadastrar_fornecedor()

    st.markdown("---")

    query_forn = """
        SELECT 
            nr_contrato AS "Nr. Contrato",
            razao_social AS "Razão Social",
            nome_fantasia AS "Nome Fantasia",
            local_atendimento AS "Local de Atendimento",
            uf AS "UF",
            cpf_cnpj AS "CPF/CNPJ",
            contato AS "Contato"
        FROM fornecedores
        ORDER BY id DESC
    """
    with get_connection() as conn:
        df_fornecedores = pd.read_sql_query(query_forn, conn)

    st.metric("Total de Fornecedores Cadastrados", len(df_fornecedores))

    # Tabela com busca integrada do Streamlit
    st.dataframe(df_fornecedores, use_container_width=True, hide_index=True)

# 7. EXCLUIR CARGAS
elif menu_selecionado == "Excluir Cargas":
    st.title("🗑️ Excluir Cargas")
    
    with get_connection() as conn:
        cargas_df = pd.read_sql_query("SELECT id, numero_carga, nome_motorista, status FROM cargas", conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga cadastrada.")
    else:
        opcoes_cargas = {f"{row['numero_carga']} - {row['nome_motorista']} ({row['status']})": row['id'] for _, row in cargas_df.iterrows()}
        selecionada = st.selectbox("Selecione a Carga a ser excluída*", list(opcoes_cargas.keys()))
        carga_id = opcoes_cargas[selecionada]

        if st.button("❌ Excluir Definitivamente", type="primary"):
            with get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM carga_administracao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_operacional WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_expedicao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM cargas WHERE id = %s", (carga_id,))
                    conn.commit()
            st.success("Carga excluída com sucesso!")
            st.rerun()

# 8. USUÁRIOS
elif menu_selecionado == "Usuários":
    st.title("👥 Gestão de Usuários")
    
    with get_connection() as conn:
        users_df = pd.read_sql_query("SELECT id, usuario, nome, perfil FROM usuarios", conn)
    st.dataframe(users_df, use_container_width=True)

    st.subheader("➕ Adicionar Novo Usuário")
    with st.form("form_novo_usuario"):
        u1, u2 = st.columns(2)
        novo_usr = u1.text_input("Usuário*").strip().lower()
        novo_nome = u2.text_input("Nome Completo*").strip().upper()

        p1, p2 = st.columns(2)
        nova_senha = p1.text_input("Senha*", type="password")
        novo_perfil = p2.selectbox("Perfil*", ["ADMIN", "PROGRAMACAO", "EXPEDICAO", "OPERACIONAL", "ADMINISTRATIVO"])

        salvar_usr = st.form_submit_button("💾 Salvar Usuário", use_container_width=True, type="primary")

        if salvar_usr:
            if not (novo_usr and novo_nome and nova_senha and novo_perfil):
                st.error("Preencha todos os campos obrigatórios.")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute(
                                "INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (%s, %s, %s, %s)",
                                (novo_usr, hash_senha(nova_senha), novo_nome, novo_perfil)
                            )
                            conn.commit()
                    st.success(f"Usuário {novo_usr} cadastrado com sucesso!")
                    st.rerun()
                except psycopg2.IntegrityError:
                    st.error("Nome de usuário já cadastrado.")
                except Exception as e:
                    st.error(f"Erro ao salvar usuário: {e}")
