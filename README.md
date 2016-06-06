ADEO PowerShell Remote Administrator Tool
==================================================
NetCat üzerinden, kullanıcı denetimini yok sayarak tam sistem yetkisiyle hedef makina üzerinde eş zamanlı komut çalıştırılmasını sağlayan araçtır.

### Programming Language

Builder: Visual Basic 
Stub: AutoIT 
Connector Script: PowerShell

### Builder

![Resim Açıklaması](http://i.hizliresim.com/pPBW3a.png)

1. Çalışacak powershell script'in bağlanacağı adres
2. Bağlantı Portu
3. Bağlantı Türü
	Reverse Connection = Tersine bağlantıdır, siz hedef makinaya değil; hedef makina size bağlanır.
	Bind Connection = Hedef makina girilen port'u dinlemeye alır ve siz hedef makinaya bağlanırsınız.
4. Kullanıcı Denetiminin(UAC) bypasslanmasında kullanılacak method türü
5. Hedef Sistemde çalıştırılacak olan exe nin yapısı
6. "Keylogger" yani klavyede basılacak tuşların dinlenmesini sağlayan fonksiyonun seçimi

### Usage

"ConnectorClient.exe" ile öncelikle kontrolü sağlayacak makinanın ip ve port bilgilerini girmek gerekmektedir.Girilen bu bilgiler stub üzerine yazılacak, ve açıldığında bu ip ve port bilgilerine netcat üzerinden bağlantı gönderecektir. 

* Not: Oluşturulan exe yi hedef sistemde çalıştırmadan önce NetCat ile girilen port'u dinlemeye almamız gerekmektedir. 

Bunun için;

```
$ Nc -lvp 443
```

Komutu kullanılır. Hedef sistemden gelen ve komut çalıştırılmaya hazır konsol ekranı aşağıdaki gibidir.

![alt tag](http://i.hizliresim.com/WY1QJP.png)

* Builder esnasında Keylogger aktifleştirilmiş ise, bağlantı aldıktan sonra "Start-KeyLogger" komutu ile klavye dinlemesi aktif hale getirilir. O andan itibaren basılan tüm tuşlar %TEMP% dizini altında keylogger.txt adıyla kaydedilecektir.

### Code Example

Stub dosyasının ilk açıldığında yaptığı işlev örneği alttaki gibidir;

```vbnet
If MeAdmin? = YES Then
	PowerShell Connect scriptini çalıştır
Else
	Beni admin yap
End IF
```

### Compile

Proje tüm haliyle açık kaynaktır. Yeniden düzenlenip derlenmesi için Stub'ın AutoIT ile compile edilmesi gerekmektedir.
AutoIT kurulum & derleme için web sitemizden bilgi edinebilirsiniz;
[Zararlı Yazılım (Malware) Geliştirmeye Giriş](http://www.adeosecurity.com/blog/siber-guvenlik/zararli-yazilim-malware-gelistirmeye-giris)

### System Requirements
* x86-32/x64 Windows 7/8/8.1/10
* Windows PowerShell v1.0

### References
* https://github.com/hfiref0x/UACME

### Authors
* [Eyüp ÇELİK](http://eyupcelik.com.tr)
* [Tolga SEZER](http://www.tolgasezer.com.tr)