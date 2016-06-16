ADEO RATShell v1.0 ~ Beta
==================================================
Powershell gücünü kullanarak hedef makina üzerinde tam yetkiye sahip olup, eş zamanlı olarak kontrol edilmesini sağlayan araçtır.

### Programming Language

Builder: Visual Basic
Stub: AutoIT
Connector Script: PowerShell

### Control Client

![1](http://i.hizliresim.com/7vNPVr.png)
![2](http://i.hizliresim.com/l1W9Lp.png)
![3](http://i.hizliresim.com/goWynQ.png)
![4](http://i.hizliresim.com/Gz05vv.png)
![5](http://i.hizliresim.com/nrW2Aa.png)
![6](http://i.hizliresim.com/l1W9kl.png)
![7](http://i.hizliresim.com/dbG05Z.png)
![8](http://i.hizliresim.com/mLk5Py.png)
![9](http://i.hizliresim.com/9LdE59.png)

### Usage & Info

Öncelikle Aynı dizindeki codes.txt, mimi_x86.txt ve mimi_x64.txt dosyalarını kendi sunucunuza/hostunuza upload ediniz.
Hedef makinada çalıştırılan exe, ilk olarak admin yetkilerini devralıp, codes.txt script'ini hiç bir yere yazmadan çalıştırmaktadır.

Kaynak kodlarını derledikten sonra açılan client ekranından "Create Server" bölümüne gelerek, bağlantı kurulacak olan makinanın ip/port bilgilerini giriniz.
Önceden upload ettiğiniz codes.txt, mimi_x86.txt ve mimi_x64.txt dosyalarının linklerini gerekli bölümlere yazınız. Daha sonra OS Arch ve UAC Bypass method'unu seçtikten sonra "create Server" diyerek, hedef makinada çalıştırılacak olan exe yi build ediniz.

Exe, Control Client ile aynı dizinde Connector.exe olarak oluşturulacaktır.

Not: Client'in ilk çalıştırıldığında girilen port, dinleme portudur. Oluşturulacak exe ye yazılan port ile aynı olmalıdır.
Not2: Client'in çalışma esnasında kütüphane hatası vermemesi için CTLs klasöründeki Registrator.exe aracı ile gerekli OCX'leri sisteme kayıt ediniz.

### Code Example

Stub dosyasının ilk açıldığında yaptığı işlev örneği alttaki gibidir;

```vbnet
If MeAdmin? = YES Then
	PowerShell Connect scriptini çalıştır
Else
	Beni admin yap ve script'i çalıştır
End IF
```

### Compile

Proje tüm haliyle açık kaynaktır. Yeniden düzenlenip derlenmesi için Stub'ın AutoIT ile compile edilmesi gerekmektedir.
AutoIT kurulum & derleme için web sitemizden bilgi edinebilirsiniz;
[AutoIT Setup & Compile](http://www.adeosecurity.com/blog/siber-guvenlik/zararli-yazilim-malware-gelistirmeye-giris)

Tüm Powershell fonksiyonları codes.txt içerisindedir. Fonksiyonları silebilir, yeni fonksiyon ekleyebilir veya devre dışı bırakabilrisiniz.
* Not: Powershell script'ini geliştirmek için Windows'un kendi derleyicisi; Powershell_ISE, alternatif olarak ise PowerGui kullanılabilir.

### System Requirements
* x86-32/x64 Windows 7/8/8.1/10
* Windows PowerShell v1.0

### Authors
* [Tolga SEZER](http://www.tolgasezer.com.tr)
* [Eyüp ÇELİK](http://eyupcelik.com.tr)

### References
* https://github.com/hfiref0x/UACME
* [Some Modules from PowerShell Empire](https://github.com/PowerShellEmpire/Empire)
* [Mimikatz - Benjamin DELPY](https://github.com/gentilkiwi/mimikatz)