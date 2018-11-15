Browser("Advantage Shopping").Page("Advantage Shopping").Sync @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Advantage Shopping").Navigate "https://collprj.abbonamento.sky.it/newaol/abbonationline?codPromo=6797&p=11100" @@ hightlight id_;_66662_;_script infofile_;_ZIP::ssf2.xml_;_

' Checking the price in the landing page
reporter.ReportEvent micDone, "Landing", "Checking the correct landing"
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("prezzoOfferta").WaitProperty "visible", True, 60000
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("prezzoOfferta").Check CheckPoint("checkPrezzoOffertaLanding") @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("14,50€")_;_script infofile_;_ZIP::ssf3.xml_;_

' Moving on
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("PROSEGUI").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("PROSEGUI")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").Sync
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("nome").WaitProperty "visible", True, 30000

' Filling data
reporter.ReportEvent micDone, "Filling", "Going to fill data"
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("nome").Set "Test" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("nome")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("cognome").Set "Microfocus" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("cognome")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("email1").Set "octane.test2018@gmail.com" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("email1")_;_script infofile_;_ZIP::ssf7.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("email2").Set "octane.test2018@gmail.com" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("email2")_;_script infofile_;_ZIP::ssf8.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("telefono").Set "3491231234" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("telefono")_;_script infofile_;_ZIP::ssf9.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("cf").Set "MCRTST80A01F205R" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("comune")_;_script infofile_;_ZIP::ssf11.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("comune").Set "MILANO" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("comune")_;_script infofile_;_ZIP::ssf12.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("indirizzo").Set "via Pietro Nenni" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("indirizzo")_;_script infofile_;_ZIP::ssf13.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("civico").Set "4" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("civico")_;_script infofile_;_ZIP::ssf14.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("cap").Set "20128" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("cap")_;_script infofile_;_ZIP::ssf15.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("provincia").Set "MI" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebEdit("provincia")_;_script infofile_;_ZIP::ssf16.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement")_;_script infofile_;_ZIP::ssf17.xml_;_

' Privacy filling
reporter.ReportEvent micDone, "Privacy", "Specifying Privacy preferences"
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("info-privacy").Select "Y" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("info-privacy")_;_script infofile_;_ZIP::ssf18.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement_2").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement 2")_;_script infofile_;_ZIP::ssf19.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("commercial").Select "Y" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("commercial")_;_script infofile_;_ZIP::ssf20.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement_3").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement 3")_;_script infofile_;_ZIP::ssf21.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("commercial2").Select "Y" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("commercial2")_;_script infofile_;_ZIP::ssf22.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement_4").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement 4")_;_script infofile_;_ZIP::ssf23.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("privacy_skyqblack").Select "Y" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("privacy skyqblack")_;_script infofile_;_ZIP::ssf24.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi")_;_script infofile_;_ZIP::ssf25.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati").Sync

' Payment Method
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati_2").WebElement("WebElement").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati 2").WebElement("WebElement")_;_script infofile_;_ZIP::ssf32.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati_2").WebRadioGroup("cc").Select "#0" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati 2").WebRadioGroup("cc")_;_script infofile_;_ZIP::ssf33.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati_2").WebElement("WebElement_2").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati 2").WebElement("WebElement 2")_;_script infofile_;_ZIP::ssf34.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati_2").WebRadioGroup("checkout-pagamento-intestatari").Select "#3" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati 2").WebRadioGroup("checkout-pagamento-intestatari")_;_script infofile_;_ZIP::ssf35.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati_3").WebButton("procedi").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati 3").WebButton("procedi")_;_script infofile_;_ZIP::ssf36.xml_;_
Browser("Advantage Shopping").Page("Acquista Sky: abbonati_3").Sync

'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement_5").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement 5")_;_script infofile_;_ZIP::ssf26.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("checkout-indirizzo").Select "VIALE SARCA,235 20126 MILANO MI" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("checkout-indirizzo")_;_script infofile_;_ZIP::ssf27.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi")_;_script infofile_;_ZIP::ssf28.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement_5").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebElement("WebElement 5")_;_script infofile_;_ZIP::ssf29.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("checkout-indirizzo").Select "VIALE SARCA,235 20126 MILANO MI" @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebRadioGroup("checkout-indirizzo")_;_script infofile_;_ZIP::ssf30.xml_;_
'Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Acquista Sky: abbonati").WebButton("procedi")_;_script infofile_;_ZIP::ssf31.xml_;_
