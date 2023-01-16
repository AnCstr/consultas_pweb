#  Modulo para consulta e extração de dados da pagina Processys Lote/guia
import os
import time
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException
from baserpa import base_selenium, base_excell
from typing import Dict

ambientes: Dict[str, str] = {"hml": "https://hml-petrobras.processys.com.br/",
                             "prd": "https://processys.saudepetrobras.com.br/"
                             }


class Consultas:
    def __init__(self, usr: str, pwd: str) -> None:
        self.__usr = usr
        self.__pwd = pwd

    def login(self, pwb: base_selenium.NavegadorWeb) -> None:
        """
        Realiza Login em Site Processys
        :return:
        """
        """username = input("Insira Usuario Processys: ")
        pwd = getpass("Informe Senha Processys: ")"""
        username = "p-Alan.castro"
        pwd = "Alca0002*"
        pwb.inserir_texto("//input[@id='username']", texto=username)
        pwb.inserir_texto("//input[@id='password']", texto=pwd)
        pwb.send_enter("//input[@id='username']")

        if not self.validador_login(pwb):
            print("Login Processys Falhou")
            raise ConnectionRefusedError

    @staticmethod
    def validador_login(pwb: base_selenium.NavegadorWeb) -> bool:
        """
        Valida se login Pweb foi bem sucedido. Retorna False caso login falhe.
        :return: True -> Caso Login Sucesso || False -> Login Falhou
        """
        try:
            pwb.clique_xpath("//a[@href='/ProcessUtilisWebService/app/view/movimentoOperacional/home/']")
            return True
        except BaseException:
            return False

    def guia_prest_protocolo(self):
        formulario = "https://processys.saudepetrobras.com.br/ProcessUtilisWebService" \
                     "/app/view/movimentoOperacional/consultas/porLoteConta/"

        pweb = base_selenium.NavegadorWeb(ambientes["prd"])
        self.login(pweb)
        pweb.navega_url(formulario)

        exl = base_excell.Excell("c:\\temp\\unicooper.xlsb", "Planilha1")

        for n in range(2, exl.ult_lin - 1):
            os.system("cls")
            print(f"Linha {n}/{exl.ult_lin}...")
            if str(exl.ws.Cells(n, 5).value) != "None":
                continue

            g_op = exl.ws.Cells(n, 1).value
            g_op = g_op.strip()

            pweb.clique_xpath("//input[@id='porLoteConta_btnLimpar']")
            pweb.inserir_texto("//input[@id='porLoteConta_numeroConta']", texto=g_op)
            pweb.duplo_clique_xpath("//input[@id='porLote_btnBuscarConta']")

            while True:
                if pweb.elemento_existe("//td[@aria-describedby="
                                        "'porLoteConta_grid_numeroGuiaPrestadorFmt']"):
                    break
                else:
                    time.sleep(0.5)
                    pweb.duplo_clique_xpath("//input[@id='porLote_btnBuscarConta']")

            """
            #  Numero Guia Prestador
            exl.ws.Cells(n, 2).value = pweb.retorna_innertext_xpath("//td[@aria-describedby="
                                                                    "'porLoteConta_grid_numeroGuiaPrestadorFmt']",
                                                                    txt=True)
            #  Numero Protocolo
            exl.ws.Cells(n, 3).value = pweb.retorna_innertext_xpath("//td[@aria-describedby="
                                                                    "'porLoteConta_grid_numprotocolo']", txt=True)
            """
            #  Valor Informado
            exl.ws.Cells(n, 5).value = pweb.retorna_innertext_xpath("//td[@aria-describedby="
                                                                    "'porLoteConta_grid_valorInformadoGuia']",
                                                                    txt=True)
            #  Valor Glosa
            exl.ws.Cells(n, 6).value = pweb.retorna_innertext_xpath("//td[@aria-describedby="
                                                                    "'porLoteConta_grid_valorGlosa']", txt=True)

        exl.fechar()
        pweb.fechar_navegador()
        print("Consulta Finalizada!")

    def get_gop_via_lote_gprest(self):
        formulario = "https://processys.saudepetrobras.com.br/ProcessUtilisWebService" \
                     "/app/view/movimentoOperacional/consultas/porLoteConta/"

        pweb = base_selenium.NavegadorWeb(ambientes["prd"])
        self.login(pweb)
        pweb.navega_url(formulario)
        n_plan = input("Informe nome da planilha")
        aba = input("Informe nome aba")

        exl = base_excell.Excell(f"c:\\temp\\{n_plan}", aba)

        for n in range(4, exl.ult_lin - 1):
            os.system("cls")
            print(f"Linha {n}/{exl.ult_lin}...")
            if str(exl.ws.Cells(n, 2).value) != "None":
                continue

            """try:
                lote = str(int(exl.ws.Cells(n, 10).value)).strip()
            except BaseException:
                exl.ws.Cells(n, 25).value = "Lote invalido"
                continue"""

            try:
                g_prest = str(int(exl.ws.Cells(n, 1).value)).strip()
            except BaseException:
                exl.ws.Cells(n, 2).value = "Guia Prestador invalida"
                continue

            g_prest = (20 - len(g_prest)) * "0" + g_prest
            pj = "03288517000116"
            #  exl.ws.Cells(n, 26).value = "Filtro Unico - Guia Prestador"
            """try:
                prt = str(int(exl.ws.Cells(n, 1).value)).strip()
            except BaseException:
                exl.ws.Cells(n, 2).value = "Protocolo invalido"
                continue"""

            try:
                pweb.clique_xpath("//input[@id='porLoteConta_btnLimpar']")
                #  pweb.inserir_texto("//input[@id='porLoteConta_numeroProtocolo']", texto=prt)
                #  pweb.inserir_texto("//input[@id='porLoteConta_numeroLote']", texto=lote)
                #  porLoteConta_numeroContratoPrestPagto
                pweb.inserir_texto("//input[@id='porLoteConta_numeroContratoPrestPagto']", texto=pj)
                while not pweb.carregou(xpath="//div[@id='divCarregandoDialog']"):
                   time.sleep(0.1)

                pweb.inserir_texto("//input[@id='porLoteConta_numeroContaPrestador']", texto=g_prest)
                pweb.duplo_clique_xpath("//input[@id='porLote_btnBuscarConta']")

                while not pweb.carregou(xpath="//div[@id='divCarregandoDialog']"):
                    time.sleep(0.1)

                alerta = pweb.verifica_alerta(xpath="//div[@id='defaultAttention']")
                if alerta != "sem_alerta":
                    if str(alerta) != "None":
                        exl.ws.Cells(n, 2).value = str(alerta)
                        continue

                try:
                    pweb.clique_xpath(xpath="//span[text()='Ok']")
                except BaseException:
                    pass
                count = 0
                if pweb.carregou(xpath="//td[@aria-describedby='porLoteConta_grid_numeroGuiaPrestadorFmt']"):
                    exl.ws.Cells(n, 2).value = "n_loc"
                    continue
                try:
                    guia_op = pweb.retorna_innertext_xpath(
                        "//td[@aria-describedby='porLoteConta_grid_numeroGuiaSenha']",
                        txt=True)
                    lote_ret = pweb.retorna_innertext_xpath(
                        "//td[@aria-describedby='porLoteConta_grid_codigoLote']",
                        txt=True)
                    carteira = pweb.retorna_innertext_xpath(
                        "//td[@aria-describedby='porLoteConta_grid_carteirinha']",
                        txt=True)
                    benef = pweb.retorna_innertext_xpath(
                        "//td[@aria-describedby='porLoteConta_grid_nomeBeneficiario']",
                        txt=True)
                    valor = pweb.retorna_innertext_xpath(
                        "//td[@aria-describedby='porLoteConta_grid_valorInformadoGuia']",
                        txt=True)

                except TimeoutException:
                    exl.ws.Cells(n, 2).value = "n_loc"
                    continue

            except UnexpectedAlertPresentException:
                for j in range(1, 5):
                    pweb.aceita_alerta()
                pweb.navega_url(ambientes["prd"])
                if self.validador_login(pweb):
                    pweb.navega_url(formulario)
                    continue
                else:
                    self.login(pweb)
                    pweb.navega_url(formulario)
                    continue
            """except TimeoutException:
                pweb.navega_url(ambientes["prd"])
                if self.validador_login(pweb):
                    pweb.navega_url(formulario)
                    continue
                while pweb.carregou("//input[@id='username']"):
                    time.sleep(1)
                    pweb.navega_url(ambientes["prd"])

                self.login(pweb)
                pweb.navega_url(formulario)
                continue"""

            exl.ws.Cells(n, 2).value = guia_op
            exl.ws.Cells(n, 3).value = lote_ret
            exl.ws.Cells(n, 4).value = carteira
            exl.ws.Cells(n, 5).value = benef
            exl.ws.Cells(n, 6).value = valor

        exl.fechar()
        pweb.fechar_navegador()
