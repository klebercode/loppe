#!/usr/bin/env python
# -*- coding: utf-8 -*-

u"""
Desenvolvido por Kleber Soares
"""

import pygtk
import gtk
import datetime
import time
import requests
import os.path
import xlwt

from time import sleep
from os import mkdir
from BeautifulSoup import BeautifulSoup
from threading import Thread

pygtk.require('2.0')
gtk.gdk.threads_init()


class Base():
    def __init__(self):
        self.janela = gtk.Window(gtk.WINDOW_TOPLEVEL)
        self.janela.set_resizable(True)
        self.janela.set_icon_from_file("icon/loppe.ico")
        self.janela.set_title("Loppe")
        self.janela.set_size_request(250, 220)
        self.janela.set_border_width(5)
        self.janela.set_position(gtk.WIN_POS_CENTER)

        self.desc = gtk.Label()
        self.desc.set_text("Localizador de placas do DetranPE\n")

        self.logo = gtk.Image()
        self.logo.set_from_file("image/logo.png")

        filter = gtk.FileFilter()
        filter.set_name("Arquivos de texto")
        filter.add_pattern("*.txt")

        self.filechooserbutton = gtk.FileChooserButton('Selecione um arquivo')
        self.filechooserbutton.add_filter(filter)

        self.progress_bar = gtk.ProgressBar()
        self.progress_bar.set_text("0/0")

        self.botao = gtk.Button("Gerar arquivo")
        self.botao.connect('clicked', self.inicio)

        self.manual = gtk.Button("Como usar")
        self.manual.connect('clicked', self.manual_msg)

        self.contato = gtk.Button("Contato")
        self.contato.connect('clicked', self.contato_msg)

        self.caixah = gtk.HBox()
        self.caixah.pack_start(self.manual)
        self.caixah.pack_start(self.contato)

        self.caixa = gtk.VBox()
        self.caixa.pack_start(self.logo)
        self.caixa.pack_start(self.desc)
        self.caixa.pack_start(self.filechooserbutton)
        self.caixa.pack_start(self.progress_bar)
        self.caixa.pack_start(self.botao)
        self.caixa.pack_start(self.caixah)

        self.janela.add(self.caixa)
        self.janela.show_all()
        self.janela.connect("delete-event", gtk.main_quit)

        gtk.main()

    def manual_msg(self, widget):
        diag = gtk.MessageDialog(self.janela, gtk.DIALOG_MODAL,
                                 gtk.MESSAGE_INFO, gtk.BUTTONS_OK)
        diag.set_markup("Como usar\n\n"
                        "1 - Crie um arquivo 'txt' com todas as placas \
                        separadas por vírgula.\n"
                        "2 - Abra o arquivo 'loppe.exe'.\n"
                        "3 - Selecione o arquivo 'txt' que você criou, \
                        contendo as placas.\n"
                        "4 - Clique em 'Gerar arquivo' e aguarde até que \
                        o processo seja concluído.")
        diag.run()
        diag.destroy()

    def contato_msg(self, widget):
        diag = gtk.MessageDialog(self.janela, gtk.DIALOG_MODAL,
                                 gtk.MESSAGE_INFO, gtk.BUTTONS_OK)
        diag.set_markup("Contato\n\n"
                        "Empresa: OW7\n"
                        "Site: ow7.com.br\n"
                        "Desenvolvedor: Kleber Soares\n"
                        "Fone: 81 8172.9074\n"
                        "Email: kleber@ow7.com.br")
        diag.run()
        diag.destroy()

    def count_in_thread(self, maximum):
        Thread(target=self.count_up, args=(maximum,)).start()

    def count_up(self, maximum):
        if self.progress_bar.get_fraction() >= float(maximum):
            fraction = 0.0
        else:
            fraction = self.progress_bar.get_fraction() + 1 / float(maximum)

        self.progress_bar.set_fraction(fraction)

    def inicio(self, widget):
        def message(msg, model):
            if model == 1:
                diag = gtk.MessageDialog(self.janela, gtk.DIALOG_MODAL,
                                         gtk.MESSAGE_WARNING, gtk.BUTTONS_OK)
            elif model == 2:
                diag = gtk.MessageDialog(self.janela, gtk.DIALOG_MODAL,
                                         gtk.MESSAGE_INFO, gtk.BUTTONS_OK)

            diag.set_markup(msg)
            diag.run()
            diag.destroy()

        day = datetime.date.today()
        month = datetime.date.today().month
        year = datetime.date.today().year
        year_month = str(year) + "-" + str(month)
        timeString = time.strftime('%H:%M:%S')
        timeString2 = time.strftime('%H-%M')
        arq_name = str(day) + "_" + timeString2

        # key validation (off)
        # URL_KEY = 'http://www.ow7.com.br/loppe.html'
        # page_key = requests.get(URL_KEY)
        # bs_key = BeautifulSoup(page_key.content)

        # a_key = bs_key.find('span', {'id': 'MM'}).string

        # print a_key

        a_key = 'teste'
        a_url = 'http://online4.detran.pe.gov.br/'
        a_url = a_url + 'NovoSite/Detran_Veiculos/result_Consulta.aspx?placa='

        if a_key != 'teste':
            message("Contacte o administrador.\n\nKleber Soares\n"
                    "81 8172.9074\nkleber@ow7.com.br", 1)
        else:
            try:
                if not os.path.exists(year_month):
                    mkdir(year_month)

                arq = open(self.filechooserbutton.get_filename())
                str_placas = arq.read()
                # strip() serve para remover as quebras das linhas
                placas = str_placas.strip().split(",")
                placas = [x for x in placas if x]
                qtd_placas = len(placas)
                arq.close()

                i = 0
                lin = 1

                wb = xlwt.Workbook()
                ws = wb.add_sheet('Detran Pernambuco')

                ws.write(0, 0, 'PLACA')
                ws.write(0, 1, 'RESTRICAO 1')
                ws.write(0, 2, 'RESTRICAO 2')
                ws.write(0, 3, 'RESTRICAO 3')
                ws.write(0, 4, 'RESTRICAO 4')
                ws.write(0, 5, 'DATA')
                ws.write(0, 6, 'HORA')

                for placa in placas:
                    placa = placa.strip()
                    i += 1
                    self.count_in_thread(qtd_placas)
                    self.progress_bar.set_text(
                        "("+placa+") "+str(i)+"/"+str(qtd_placas))

                    while gtk.events_pending():
                        gtk.main_iteration()

                    URL_ULTIMOS_RESULTADOS = a_url + placa
                    page = requests.get(URL_ULTIMOS_RESULTADOS)
                    bs = BeautifulSoup(page.content)

                    labels = (
                        bs.find('span', {'id': 'lblRestricao1'}
                                ).find('font').string,
                        bs.find('span', {'id': 'lblRestricao2'}
                                ).find('font').string,
                        bs.find('span', {'id': 'lblRestricao3'}
                                ).find('font').string,
                        bs.find('span', {'id': 'lblRestricao4'}
                                ).find('font').string,
                    )

                    # csv = placa+","
                    ws.write(lin, 0, placa)

                    col = 1

                    for label in labels:
                        if not label:
                            ws.write(lin, col, label)
                        col += 1

                    ws.write(lin, 5, str(day))
                    ws.write(lin, 6, timeString)

                    lin += 1

                    sleep(1)

                    wb.save(year_month+"/"+arq_name+".xls")

                message("Arquivo gerado com sucesso.\n"
                        "Verifique a pasta do aplicativo.", 2)
            except TypeError, erro:
                if not self.filechooserbutton.get_filename():
                    message("Selecione um arquivo.", 1)
                else:
                    print "Um erro ocorreu: %s" % erro
                    message("Um erro ocorreu: %s" % erro, 1)

        self.progress_bar.set_fraction(0.0)
        self.progress_bar.set_text("0/0")

if __name__ == '__main__':
    Base()
