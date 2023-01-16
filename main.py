from lote_conta import Consultas

if __name__ == "__main__":
    c = Consultas("p-alan.castro", "Alca0001*")
    #  c.guia_prest_protocolo()
    c.get_gop_via_lote_gprest()
