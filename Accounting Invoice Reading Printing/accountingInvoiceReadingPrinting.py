import pytesseract
import cv2
import sys
import os
import glob
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import re
import pandas as pd

class Pencere(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.dosyaAc = QtWidgets.QPushButton("İşlem Görecek Dosya")
        v_box = QtWidgets.QVBoxLayout()
        v_box.addWidget(self.dosyaAc)
        v_box.addStretch()
        self.setLayout(v_box)
        self.dosyaAc.clicked.connect(self.islemler)
        self.show()

    def islemler(self):
        # klasörü seçtim
        dosyaAdi = QFileDialog.getExistingDirectory(self, "Klasör Seç", "C:\\")
        dosyaAdi += "/*"
        dosyaYolu = glob.glob(dosyaAdi)
        dosyaYolu_Filtre = []
        for i in dosyaYolu:
            if i.endswith("jpg") or i.endswith("jpeg") or i.endswith("png"):
                dosyaYolu_Filtre.append(i)

        dosya_yolu = QFileDialog.getSaveFileName(self, "İşlenmiş Dosyaların Kaydedildileceği Yeri Seç",
                                                 os.getenv("HOME"))
        yeni_dosya_yolu = dosya_yolu[0]
        index = yeni_dosya_yolu.rfind('/')
        dosya_yolu_isim = yeni_dosya_yolu[index + 1::]
        dosya_yolu_yolu = yeni_dosya_yolu[0:index + 1]
        path = os.path.join(dosya_yolu_yolu, dosya_yolu_isim)
        os.mkdir(path)
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        firmAdi = ""
        fisNo = ""
        topkdv = ""
        toplam = ""
        filtrelenmis_ifade=""
        firma_adi_ihtimal = ["a.ş.", "a.s.", "a.s", "ins.san.ve", "ins.san", "as."]
        fis_no_ihtimal = ["fis", "fiş", "fig", "fls", "flg", "lis"]
        kdv_ihtimal = ["topkdv", "toplamkdv"]
        toplam_fiyat_ihtimal = ["toplam", "top"]

        def sayiVarMi(kelimeler):
            for i in kelimeler:
                if i.isdigit():
                    return True
            return False


        def karsilikBulma(ihtimal_dizi, index, text_dizi):
            for l in ihtimal_dizi:
                if l == text_dizi[index]:
                    for x in range(1, 5):
                        if sayiVarMi(text_dizi[index + x]):
                            text_dizi[index + x] = text_dizi[index + x].strip("*")
                            text_dizi[index + x] = text_dizi[index + x].replace(",", ".")
                            c=text_dizi[index+x].endswith(",")
                            if c:
                                text_dizi[index+x]+=c
                                
                            return text_dizi[index + x]

        sayac = 0
        firma_isim_k = 0
        ay_sozluk = {"01": "ocak", "02": "subat", "03": "mart", "04": "nisan", "05": "mayıs", "06": "haziran",
                     "07": "temmuz", "08": "agustos", "09": "eylul", "10": "ekim", "11": "kasim", "12": "aralik",
                     "-1": "AYIKLANAMAYANLAR"}
        for y in dosyaYolu_Filtre:
            resim = cv2.imread(y)
            gray = cv2.cvtColor(resim, cv2.COLOR_BGR2GRAY)
            text = pytesseract.image_to_string(gray)
            text = text.replace("#", "*")
            text = text.replace("x", "%")
            text = text.replace("$", "ş")
            text = text.lower()
            date_extract_pattern = "(\d{2})\.(\d{2})\.(\d{4})"
            matches1 = re.findall(date_extract_pattern, text)
            if len(matches1) == 0:
                date_extract_pattern = "[0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{4}"
                matches1 = re.findall(date_extract_pattern, text)
            matches1 = str(matches1)
            tarih = matches1
            tarih = tarih.strip("[(")
            tarih = tarih.strip(")]")
            tarih = tarih.replace("'", "")
            tarih = tarih.replace(",", "/")
            tarih = tarih.replace(" ", "")
            alinan_ay = tarih[3:5]
            
            try:
                ay_sozluk[alinan_ay]
            except KeyError:
                alinan_ay = "-1"

            text = text.split()
            for z in range(0, len(text)):
                for l in firma_adi_ihtimal:
                    if l == text[z]:
                        for i in reversed(range(0, 3)):
                            firmAdi += text[z - i] + " "
                            firma_isim_k += 1
                    if firma_isim_k > 1:
                        break

                if karsilikBulma(fis_no_ihtimal, z, text) is not None:
                    fisNo = karsilikBulma(fis_no_ihtimal, z, text)

                if karsilikBulma(kdv_ihtimal, z, text) is not None:
                    topkdv = karsilikBulma(kdv_ihtimal, z, text)

                if karsilikBulma(toplam_fiyat_ihtimal, z, text) is not None:
                    toplam = karsilikBulma(toplam_fiyat_ihtimal, z, text)

            filtre_kelime = [tarih, firmAdi, fisNo, toplam, topkdv]
            for j in range(0, len(filtre_kelime)):
                if filtre_kelime[j] == "":
                    filtre_kelime[j] = "Bulunamadı"
                filtrelenmis_ifade = filtre_kelime[0] + "," + filtre_kelime[1] + "," + filtre_kelime[2] + "," + \
                                     filtre_kelime[3] + "," + filtre_kelime[4] + "\n"

            olusturulacakYol = yeni_dosya_yolu + "/" + ay_sozluk[alinan_ay] + ".csv"
            olusacakDosya = open(olusturulacakYol, "a")

            if len(glob.glob(yeni_dosya_yolu + "/*")) == sayac:
                olusacakDosya.write(filtrelenmis_ifade)
            else:
                olusacakDosya.write("Tarih,Kimden Alindigi,Fis No,Toplam Tutar,Toplam KDV\n" + filtrelenmis_ifade)
                sayac += 1

            firmAdi = ""

        csv_dosya = glob.glob(yeni_dosya_yolu + "/*")
        for i in csv_dosya:
            readFile = pd.read_csv(i, encoding='unicode_escape')
            sorted_readFile=readFile.sort_values(by='Tarih')
            yeni_dosya = i[:-3]
            sorted_readFile.to_excel(yeni_dosya + "xlsx", index=None, header=True)

        olusacakDosya.close()
        for h in csv_dosya:
            if os.path.exists(h) and os.path.isfile(h):
                os.remove(h)


app = QtWidgets.QApplication(sys.argv)
pencere = Pencere()
sys.exit(app.exec_())
