'''
    Autor:
     C. Henrique
    Descrição:
     Feito para realização de consultas no servidor Maranhão e geração de relatórios em excel.
    Versão:
     1.3v
    Build:
     Adição do loop e ajuste na query.
'''


import config.config as cfg
def buscar():
    print("Escolha um banco para Consultar :")
    print("1 . Loja F04")
    print("2 . Loja F05")
    print("3 . Loja F10")
    print("4 . Loja F11")
    print("5 . Loja F17")
    print("6 . Loja F18")
    print("7 . Loja F20")
    print("8 . Loja F22")

    opcao = int(input("Digite somente uma opção: "))
    if opcao == 1:
        db = cfg.connection('Pcicero_F04', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F04'
        print(f'Conectando na loja {loja}')
    elif opcao == 2:
        db = cfg.connection('Pcicero_F05', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F05'
        print(f'Conectando na loja {loja}')
    elif opcao == 3:
        db = cfg.connection('Pcicero_F10', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F10'
        print(f'Conectando na loja {loja}')
    elif opcao == 4:
        db = cfg.connection('Pcicero_F11', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F11'
        print(f'Conectando na loja {loja}')
    elif opcao == 5:
        db = cfg.connection('Pcicero_F17', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F17'
        print(f'Conectando na loja {loja}')
    elif opcao == 6:
        db = cfg.connection('Pcicero_F18', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F18'
        print(f'Conectando na loja {loja}')
    elif opcao == 7:
        db = cfg.connection('Pcicero_F20', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F20'
        print(f'Conectando na loja {loja}')
    elif opcao == 8:
        db = cfg.connection('Pcicero_F22', '192.168.0.78', 'consulta', 'A1b2c3@')
        loja = 'F22'
        print(f'Conectando na loja {loja}')
    else:
        print("Erro na escolha da opção!")
        input("Aperte Enter para cancelar ....")
        quit()

    codigo = int(input('Digite o codigo do produto: '))

    res = cfg.consulta(db,
                       f"SELECT ITNFCompraST.CDPRD,NFCOMPRA.NRDOC,ITNFCOMPRAST.QTATD,NFCOMPRA.DTEMS,ITNFCOMPRAST.PRUNI,ITNFCompraST.PCICM,ITNFCOMPRAST.PCIPI,"
                       f"NFCOMPRA.CDCFO,FORNECEDOR.NMFRN,FORNECEDOR.NRCGC,NFCOMPRA.CHAVE_ACESSO,ITNFCOMPRAST.CDPRD,NFCOMPRA.CDFRN,FORNECEDOR.UFFRN,NFCOMPRA.DTENT "
                       f"FROM ITNFCOMPRAST,NFCOMPRA,FORNECEDOR "
                       f"WHERE ITNFCOMPRAST.nota_id=NFCOMPRA.row_id "
                       f"AND NFCOMPRA.CDFRN=FORNECEDOR.CDFRN "
                       f"AND ITNFCompraST.CDPRD={codigo} "
                       f"AND NFCOMPRA.CDCFO IN (1403,2403) "
                       f"ORDER BY nfcompra.DTEMS desc")

    consul = cfg.list_db(res)
    cfg.import_excel(consul, codigo, loja)


buscar()

loop = int(input("Caso queira repetir digite o numero 1 para Sim e 0 para Sair :"))
if loop == 1:
    buscar()
else:
    print("Finalizando..")
    input("Aparte Enter para sair!")
    quit()
