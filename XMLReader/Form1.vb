Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Xml

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode_order As XmlNodeList
        Dim xmlnode_client As XmlNodeList
        Dim xmlnode_delivery As XmlNodeList
        Dim xmlnode_item As XmlNodeList
        Dim i As Integer
        Dim rakuten As String
        Using client = New WebClient()
            client.Encoding = Encoding.UTF8
            rakuten = client.DownloadString("http://webservice.rakuten.de/merchants/orders/getOrders?key=123456789a123456789a123456789a12")
            Console.WriteLine(rakuten)
            Dim xmlfile As System.IO.StreamWriter
            xmlfile = My.Computer.FileSystem.OpenTextFileWriter("\\SERVER-02\BMECat\Rakuten\rakuten.xml", False)     ' Write to File
            xmlfile.Write(rakuten)
            xmlfile.Close()
        End Using
        Dim fs As New FileStream("C:\Users\s.ewert\Desktop\rakuten.xml", FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        xmlnode_order = xmldoc.GetElementsByTagName("order")
        xmlnode_client = xmldoc.GetElementsByTagName("client")
        xmlnode_delivery = xmldoc.GetElementsByTagName("delivery_address")
        xmlnode_item = xmldoc.GetElementsByTagName("item")

        Console.WriteLine(xmldoc.ChildNodes)

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("\\SERVER-02\BMECat\Rakuten\rakuten.csv", False)     ' Write to File

        file.WriteLine("Bestellnummer;Bestelldatum;EMail;Artikelnummer;Menge;Preis;" +
                          "Empfangfirma;Empfangvorname;Empfangnachname;EmpfangStrasse;EmpfangPLZ;Empfangort;EmpfangLKZ;" +
                          "Rechnungfirma;Rechnungvorname;Rechnungnachname;RechnungStrasse;RechnungPLZ;Rechnungort;RechnungLKZ;" +
                          "Kundennummer;Zahlungsart;Auftragsart;Versandnummer;Endetext;Versandtext")        'Kopfzeile


        Dim f As Integer = -1
        Dim offset As Integer = 0 ' Items in first delivery
        For i = 0 To xmlnode_order.Count - 1

            Dim delivery_item As String = xmlnode_order(i).ChildNodes.Item(12).InnerXml
            Dim phrase As String = "<item>"
            offset = (delivery_item.Length - delivery_item.Replace(phrase, String.Empty).Length) / phrase.Length    ' Get number of items in delivery


            For f = f + 1 To f + offset
                Try
                    f = WriteXMLLine(xmlnode_order, xmlnode_client, xmlnode_delivery, xmlnode_item, i, file, f, False)  'Write items of delivery
                Catch ex As Exception

                End Try
            Next

            f = f - 1
            WriteXMLLine(xmlnode_order, xmlnode_client, xmlnode_delivery, xmlnode_item, i, file, f, True)   'Write shipping
        Next

        file.Close()

        Application.Exit()
    End Sub



    ' Write some variables from xml to txt document
    Private Shared Function WriteXMLLine(xmlnode_order As XmlNodeList, xmlnode_client As XmlNodeList, xmlnode_delivery As XmlNodeList, xmlnode_item As XmlNodeList, i As Integer, file As StreamWriter, f As Integer, shipping As Boolean) As Integer
        Dim artikelnummer As String = xmlnode_item(f).ChildNodes.Item(1).InnerText.Trim() 'Artikelnummer; Eventuell erneut prüfen
        Dim menge As String = xmlnode_item(f).ChildNodes.Item(6).InnerText.Trim() 'Menge
        Dim preis As String = xmlnode_item(f).ChildNodes.Item(7).InnerText.Trim() 'Preis

        If shipping Then
            artikelnummer = "VERSAND-1955_LAGER"
            menge = "1"
            preis = "3.90"
        End If

        Try
            file.WriteLine(xmlnode_order(i).ChildNodes.Item(0).InnerText.Trim() + 'Bestellnummer
                                  ";" + xmlnode_order(i).ChildNodes.Item(9).InnerText.Trim() + 'Bestelldatum
                                  ";" + xmlnode_client(i).ChildNodes.Item(11).InnerText.Trim() + 'EMail
                                  ";" + artikelnummer + 'Artikelnummer; Eventuell erneut prüfen
                                  ";" + menge + 'Menge
                                  ";" + preis + 'Preis
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(3).InnerText.Trim() + 'Empfängerfirma
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(1).InnerText.Trim() + 'Empfängervorname
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(2).InnerText.Trim() + 'Empfängernachname
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(4).InnerText.Trim() + xmlnode_delivery(i).ChildNodes.Item(5).InnerText.Trim() + 'Empfängerstraße
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(7).InnerText.Trim() + 'Empfänger PLZ
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(8).InnerText.Trim() + 'Empfängerort
                                  ";" + xmlnode_delivery(i).ChildNodes.Item(9).InnerText.Trim() + 'Empfänger LKZ
                                  ";" + xmlnode_client(i).ChildNodes.Item(4).InnerText.Trim() + 'Rechnungsfirma
                                  ";" + xmlnode_client(i).ChildNodes.Item(2).InnerText.Trim() + 'Rechnungsvorname
                                  ";" + xmlnode_client(i).ChildNodes.Item(3).InnerText.Trim() + 'Rechnungsnachname
                                  ";" + xmlnode_client(i).ChildNodes.Item(5).InnerText.Trim() + xmlnode_client(i).ChildNodes.Item(6).InnerText.Trim() + 'Rechnungsstraße
                                  ";" + xmlnode_client(i).ChildNodes.Item(8).InnerText.Trim() + 'Rechnungs PLZ
                                  ";" + xmlnode_client(i).ChildNodes.Item(9).InnerText.Trim() + 'Rechnungsort
                                  ";" + xmlnode_client(i).ChildNodes.Item(10).InnerText.Trim() + 'Rechnungs LKZ
                                  ";" + "1234567891" + 'Kundennummer
                                  ";" + "10" + 'Zahlungsart
                                  ";" + "7" + 'Auftragsart
                                  ";" + "3" + 'Versandnummer
                                  ";" + "Bitte nicht überweisen. Die Rechnung wurde bereits über Rakuten beglichen" + 'Endetext
                                  ";" + "Rakuten") 'Versandtext
        Catch ex As Exception

        End Try

        Return f
    End Function
End Class
