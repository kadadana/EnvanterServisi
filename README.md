Envanter Servisi
Bu uygulama sirketlerin bilgisayar envanter yonetiminin kolaylasmasi adina yaratilmis bir Windows Servisidir.
Servis bilgisayar calistigi surece calisir.
Bilgisayar acildigi an ve gunun belirli bir saati bilgisayarin seri numarasi, adi, modeli, 
anlik olarak acik olan kullanici vs. gibi bilgileri SQL Server uzerindeki EnvanterTablosu isimli tabloya ve
kendine ait asset numarasiyla olusturulmus tabloya yazar.

Her bilgisayarda MSSQL Server kurmanin zahmetli olacagindan sadece ortak bir API'ye veri gonderen ancak 
verileri lokal olarak SQL'e yazmayan versiyonunu urettim:
https://github.com/kadadana/EnvanterServis
