Single - jednostruka preciznost
Double - dvostruka preciznost

-----------------------------------------------------------------------------------------

5.1 (3 poena) Napisati VBA funkciju Fun koja za argument ima realan broj dvostruke preciznosti x, i koja vraća rezultat definisan sljedećim izrazom (funkcija Sqr vraća kvadratni korijen broja):
	
		Function Fun(x as Double) as Double
			If x > 0 Then
				Fun = 1/x + 1/(x^2)
			Elseif x < 0 Then
				Fun = 11^(1/x+1)
			Elseif  x = 0 Then
				Fun = sqr(6.7)
			Endif
		End Function



5.2 (3 poena) Napisati VBA funkciju Fun koja za argument ima realan broj dvostruke preciznosti, i koja vraća rezultat definisan sljedećim izrazom:
 
Function Fun(K as Double) as Double
	Dim S as Integer
	If x < -3 Then
		Fun = 2/(1-x)^2
	Elseif x = -3 Then
		Fun = 3^x
	Elseif X > -3 Then
		For K = 1 to 155 Step 1
			S =  S + K
		Next
	Endif
End Function












5.3 (3 poena) Napisati VBA funkciju Fun koja za argumente ima realan broj dvostruke preciznosti x i cio broj N, i koja vraća rezultat definisan sljedećim izrazom (pretpostaviti da je N pozitivno):
Function Fun(x as Double, N as Integer) as Double
	Dim S as Integer
	If x < 2 Then
		Fun = 2/(1-x^2)
	Elseif X = 2 Then
		For k = 1 to N Step 1
			S = S + k
		Next
		Fun = S
	Elseif x > 2 Then
		Fun = 3^(1/N)
	Endif
End Function



5.4 (3 poena) Napisati VBA funkciju Fun koja za argumente ima realan broj dvostruke preciznosti x i cio broj K, i koja vraća rezultat definisan sljedećim izrazom (pretpostaviti da je K pozitivno):

	Function Fun(x as Double, K as Integer) as Double
		Dim S1 as Integer, S2 as Integer
		If x > 2 Then
			For n=1 to K Step 1
				S = S + (n^2)
			Next
			Fun = S
		Elseif x = 2 Then
			Fun = 1/(2+(x^2))
		Elseif x < 2 Then
			Fun = 1+2x+3(x^2)+4(x^3)

















5.5 (3 poena) Napisati VBA funkciju ParNepar koja za argument ima cijeli broj N, i koja vraća rezultat definisan sljedećim izrazom:

	Function ParNepar(N as Integer) as Integer
		Dim S1 as Integer, S2 as Integer, k as integer
		If N Mod 2 = 1 Then
			For k=1 to N Step 2
				S1 = S1 + k^2
			Next
			ParNepr = S1
		Elseif N Mod 2 = 0 Then
			For k=2 to N Step 2
				S2 = S2 + k^2
			Next
			ParNepr = S2
		Endif
	End Function



5.6 (3 poena) Napisati VBA funkciju Izraz koja za argument ima cijeli broj N, i koja vraća rezultat definisan sljedećim izrazom:

	Function Izraz(N as Integer) as Integere
		If N >= -2 And N < 4 Then
			Izraz = (N-2)^3
		Elseif N >= -10 And N < -2
			Izraz = 4-N^2
		Elseif N < -10 Or N >= 4
			Izaz = 97
		Endif
	End Function



















_________________________________________________________________________________________

6.1 (4 poena) Napisati VBA funkciju Fun1 koja za argumente ima String S. Funkcija treba da vrati string koji se dobija samo od malih suglasnika stringa S.

Primjer: Fun1(“Samo Odgovorno”) vraca string “mdgvrn”

	Function Fun1(S as String) as String
		Dim S1 as String, I as Integer
		S1 = “”
		For I=1 to Len(S)
			If Mid(S,I,1) Like “[bcdgfhjklmn]” Then
				S1 = S1 & Mid(S,I,1)
			End If
		Next
		Fun1 = S1
	End Function

6.2 (4 poena) Napisati VBA funkciju Fun2 koja za argument ima string S. Funkcija treba da vrati string koji se dobija od karaktera stringa S koji nisu ni slova ni cifre. 

Primjer: Fun2(“ABC abc$%%.9”) vraca sring “$%%”

	Function Fun2(S as String) as String
Dim S1 as String, I as Integer
		S1 = “”
		For I = 1 to Len(S)
			If Not Mid(S,I,1) Like “[a-z]” Or Mid(S,I,1) Like “[A-Z]” Or Mid(S,I,1) 
Like “[0-9]” Then
				S1 = S1 & Mid(S,I,1)
			End If
		Next
		Fun2 = S1
	End Function

6.3 (4 poena) Napisati VBA funckciju Eliminacija koja za argumente ima string S i dva karaktera K1 i K2. Funkcija treba da vrati string koji se dobija kad se iz stringa S izbace karakteri K1 I K2. 

Primjer: Izbaci(“Odgovornost”,”o”,”s”) vraca string “Odgvrnt”

	Function Eliminacija(S as String, K1 as String, K2 as String) as String
Dim S1 as String, I as Integer
		S1 = “”
		For I = 1 to Len(S)
			If Not Mind(S,I,1) Like K1 And Not Mind(S,I,1) Like K2
				S1 = S1 & Mid(S,I,1)
			End If
		Next
		Eliminacija = S1
	End Function
6.4 (4 poena) Napisati VBA funkciju Krajevi koja za argument ima string S i koja vraca string dobijen tako sto se stringu S zamijeni prvi i posljednji karakter, 

Primjer: Krajevi(“12345#abc”) vraca string “c2345#ab1”

	Function Krajevi(S as String) as String
		Dim S1 as String
			S1 = Right(S,1) & Mid(S,2,Len(S)-2) & Left(S,1)
		Krajevi = S1
	End Function
	
6.5 (4 poena) Napisati VBA funkciju IzbaciDolar koja za argument ima string S i 
koja vraća string koji se dobija kad se iz stringa S izbaci svaka pojava karaktera $. 

Primjer: IzbaciDolar(“123$45#abc”) vraća string “12345#abc”.

	Function IzbaciDolar(S as String) as String
		Dim I As Integer, S1 As String
   		S1 = ""
    		For I = 1 To Len(S)
    		If Mid(S, I, 1) Like "[!$]" Then     ''' If Not Mid(S, I, 1) Like "$" Then
        		S1 = S1 + Mid(S, I, 1)
    		End If
    		Next
    		IzbaciDolar = S1
	End Function
		

6.6 (4 poena) Napisati VBA funkciju DolarZvijezda koja za argument ima string S i koja svaku pojavu karaktera $ u stringu S mijenja karakterom *. 

Primjer: DolarZvijezda(“123$45#abc”) vraća string “123*45#abc”.

	Function DolarZvijezda()
   		Dim I As Integer
    		S = "Da$n$ijela_ikona$$"
    		For I = 1 To Len(S)
    		If Mid(S, I, 1) Like "$" Then
      	 	Mid(S, I, 1) = "*"
    		End If
    		Next
  		DolarZvijezda = S
	End Function






6.7 (4 poena) Napisati VBA funkciju UmanjiElemente koja za argumente ima niz Double brojeva X i Double broj K. Funkcija treba da umanji za 1 svaki element niza X koji je veći od broja K. Funkcija vraća broj ovakvih umanjenja.

	Function UmanjiElemente(X() as Double, K as Double) as Integer
		Dim I as Integere, BR as Integer
		For I = LBound(X) to UBound(X)
		If X(I) > K Then
			X(I) = X(I)-1
			BR = BR+1
		Endif
		Next I
		UmanjiElemente = BR
	End Function

6.8 (4 poena) Napisati VBA funkciju UvećajElemente koja za argumente ima niz Double brojeva X i Double broj Y. Funkcija treba da uveća za Y svaki element niza X koji je veći od Y. Funkcija vraća broj ovakvih uvećanja.


	Function UvecajElemente(X() as Double, Y as Double) as Integer
		Dim I as Integere, BR as Integer
		For I = LBound(X) to UBound(X)
		If X(I) > Y Then
			X(I) = X(I)+Y
			BR = BR+1
		Endif
		Next I
		UvecajElemente = BR
	End Function

6.9 (4 poena) Napisati VBA funkciju PonoviString koja za argumente ima string S i cio broj N. Funkcija treba da vrati string dobijen ponavljanjem stringa S tačno N puta. Ukoliko N nije pozitivan broj, funkcija treba da vrati prazan string. 

Primjer: PonoviString(“abc”,3) vraća string “abcabcabc”, dok PonoviString(“abc”,-2) vraća string “”.

	Function PonoviString(S as String, N as Integere) as String
		Dim I as Integer
		S1 = ""
		If N > 0 Then
			For I = 1 To N
				S1 = S1 + S
			Next
		Elseif N < 0 Then
			S1 = “”
		EndIf
		PonoviString = S1	
	End Function
6.10 (4 poena) Napisati VBA funkciju Prebacivanje koja za argumente ima string S i cio broj N. Funkcija vraća string koji se dobija kad se posljednjih N karaktera stringa S prebace na njegov početak. Ukoliko je N veće od dužine stringa ili je manje od 1, funkcija treba da vrati string S. 

Primjer: Prebacivanje(“abcd”,3) vraća string “bcda”, dok Prebacivanje(“abcd”,6) vraća string “abcd”.

	Function Prebacivanje(S as String, N as Integer) as String
		S1 = ""
		If N > Len(S) or N < 1 Then 
			Prebacivanje = S
		Else
			S = Right(S,N) & Left(S,N-2)
		Endif
	End Function


6.11 (4 poena) Napisati VBA funkciju PrebaciCifru koja za argument ima string S i koja vraća string koji se dobija kad se prvi karakter stringa S prebaci na njegov kraj, samo pod uslovom da je prvi karakter cifra. U suprotnom, funkcija vraća string S. 

Primjer: PrebaciCifru(“12345abc”) vraća string “2345abc1”, dok PrebaciCifru(“abc”) vraća “abc”.

	Function PrebaciCifru(S as String) as String
		If Mid(S,1,1) Like "[0-9]" Then
			S = Mid(S,2) & Mid(S,1,1)
		Else
			PrebaciCifru = S
		Endif
		PrebaciCifru = S
	End Function

6.12 (4 poena) Napisati VBA funkciju Fun3 koja za argument ima string S. Funkcija treba da vrati broj karaktera stringa S koji nisu cifre. Primjer: Fun3("Kolokvijum 2018.") vraća broj 12.

	Function Fun3(S as String) as Integer
		Dim BR as Integer
		For I=1 To Len(S)
			If Not Mid(S,I,1) Like "[0-9]" Then
				BR = BR+1
			End If
		Next
		Fun3 = BR
	End Function



6.13 (4 poena) Napisati VBA funkciju VelikiSusjedi koja za argument ima string S. Funkcija treba da odredi i vrati koliko puta se u stringu nalaze dva velika slova jedno pored drugog. 

Primjer: VelikiSusjedi("VBA 2018 Kolokvijum") vraća broj 2 (susjedi V i B, B i A).

	Function VelikiSusjedi(S as String) as Integer
		Dim BR as Integer, I as Integer
		For I=1 To Len(S)
			If Mid(S,I,1) Like "[A-Z]" And Mid(S,I+1,1) Like "[A-Z]" Then
				BR=BR+1
			End If
		Next
		VelikiSusjedi = BR
	End Function


6.14 (4 poena) Napisati VBA funkciju PocetniSamoglasnici koja za argument ima string S, i koja vraća broj uzastopnih početnih karaktera koji su mali samoglasnici (a,e,i,o,u).

Primjer: Funkcija PocetniSamoglasnici("aeokDEF123") treba da vrati broj 3.

	Function PocetniSamoglasnici(S as String) as Integer
		Dim BR as Integer, I as Integer
		For I=1 To Len(S)
			If Mid(S,I,1) Like "[aeiou]" Then 
				BR = BR+1
			ElseIf Not Mid(S,I,1) Like "[aeiou]" Then Exit For
			End If
		Next
		PocetniSamoglasnici = BR
	End Function
		


















6.15 (4 poena) Napisati VBA funkciju koja za argument ima dva stringa P i Q. Funkcija treba da vrati onaj string koji ima više velikih slova P, Q, R, S i T.

	Function VelikaSlova(P as String, Q as String) as String
		Dim BR1 as Integer, BR2 as Integer
		For I=1 To Len(P)
			If Mid(P,I,1) Like "[PQRST]" Then
				BR1 = BR1 + 1
			Endif
		Next
		For I=1 To Len(Q)
			If Mid(Q,I,1) Like "[PQRST]" Then
				BR2 = BR2 + 1
			Endif
		Next
		If BR1 > BR2 Then
			VelikaSlova = P
		Elseif BR1 < BR2
			VelikaSlova = Q
		End If
	End Function
	
	
6.16 (4 poena) Napisati VBA funkciju ViseMalih koja za argument ima dva stringa S i T. 
Funkcija treba da vrati posljednji karakter onog stringa koji ima više malih slova.

	Function ViseMalih(S as String, T as String) as String
		Dim BR1 as Integer, BR2 as Integer
		For I=1 To Len(S)
			If Mid(S,I,1) Like "[a-z]" Then
				BR1 = BR1 + 1
			Endif
		Next
		For I=1 To Len(T)
			If Mid(T,I,1) Like "[a-z]" Then
				BR2 = BR2 + 1
			Endif
		Next
		If BR1 > BR2 Then
			ViseMalih = Right(S,1) 
		Elseif BR1 < BR2 Then
			ViseMalih = Right(T,1)
		End If
	End Function





6.17 (4 poena) Napisati VBA funkciju BarDuplo koja za argumente ima niz Double brojeva X i jedan Double broj A. Funkcija treba da vrati broj elemenata niza X koji su bar duplo veći od A.

	Function BarDuplo(X() as Double, A as Double) as Integer
		Dim I as Integere, BR as Integer
		BR = 0
		For I = LBound(X) to UBound(X)     ''' uvijek kad je niz   
			If X(I) > 2*A Then
				BR = BR + 1
			Endif
		Next
		BarDuplo = BR
	End Function



6.18 (4 poena) Napisati VBA funkciju Izmedju koja za argumente ima Integer niz W i Integer brojeve X i Y, pri čemu je X>Y, što ne treba provjeravati. Funkcija treba da u svaki element niza koji je po vrijednosti između X i Y upiše 0. Funkcija vraća broj izvršenih zamjena.

	Function Izmedju(W() as Integer, X as Integer, Y as Integer) as Integer
		Dim BR as Integer, I as Integer
		BR = 0
		For I = LBound(W) to UBound(W) 
			If W(I) < X And W(I) > Y Then
				W(I) = 0
				BR = BR+1
		Endif
		Next
		Izmedju = BR
	End Function

6.19 (4 poena) Napisati VBA funkciju PozitivniMax2 koja za argumente ima niz Double brojeva Niz i jedan Double broj X. Funkcija treba da vrati broj elemenata niza Niz koji su pozitivni i nisu veći od dvostruke vrijednosti broja X.

	Function PozitivniMax2(Niz() as Double, X as Double) as Integer
		Dim BR as Integer, I as Integer
		For I = LBound(Niz) to UBound(Niz)
			If Niz(I) > 0 And Niz(i) < 2X
				BR = BR + 1
			Endif
		Next
		PozitivniMax2 = BR
	End Function





6.20 (4 poena) Napisati VBA funkciju SredinaPozitivni koja za argument ima Single niz X. 
Funkcija treba da vrati aritmetičku sredinu svih pozitivnih elemenata niza.

	Function SredinaPozitivna(X() as Single) as Single
		Dim I as Integere, BR as Integere, S as Integer, A as Single
		For I = LBound(X) to UBound(X)
			If X(I) > 0 Then
				S = S + X(I)
				BR = BR + 1
			Endif
		Next
		A = S/BR
		SredinaPozitivna = A
	End Function





6.21 (4 poena) Napisati VBA funkciju UmanjiElemente koja za argumente ima niz Double brojeva X i Double broj K. Funkcija treba da umanji za 1 svaki element niza X koji je veći od broja K. Funkcija vraća broj ovakvih umanjenja.


	Function UmanjiElemente(X() as Double, K as Double) as Integr
		Dim BR as Integer
		For I = LBound(X) To LBound(X)
			If X(I) > K Then
				X(I) = X(I) – 1
				BR = BR + 1
			End If
		Next
		UmanjiElemente = BR
	End Function











































7.1 (4 poena) Napisati VBA funkciju PrviVeciSusjedi koja za argument ima niz cijelih brojeva X i vraća prvi element tog niza koji je veći od svojih susjeda (prvog sa lijeva i prvog sa desna). Prvi i zadnji element niza ne uzimati u obzir jer imaju samo jednog susjeda. Ukoliko ne postoji nijedan takav element, vratiti posljednji element niza.

Primjer: Ako se funkciji proslijedi niz [7, 3, 2, 8, 4, 7, 1], ona treba da vrati broj 8.

	Function PrviVeciSusjedi(X() as Integer) as Integer
		Dim IND as Integer
		For I = LBound(X)+1 to UBound(X)-1
			If X(I) > X(I+1) And X(I) > X(I-1) Then
				PrviVeciSusjedi = X(I)
				IND = 1
				Exit For
			End If
		Next
		If IND = 0 Then
			PrviVeciSusjedu = X(UBound(X))
		End If
	EndFunction
	
7.2 (4 poena) Napisati VBA funkciju ManjiSusjedi koja za argument ima niz realnih brojeva X i vraća posljednji element tog niza koji je manji od svojih susjeda (prvog sa lijeva i prvog sa desna). Prvi i zadnji element ne uzimati u obzir jer imaju samo jednog susjeda. Ukoliko ne postoji nijedan takav element, vratiti prvi element niza.

Primjer: Ako se funkciji proslijedi niz [6.9, 1.1, 3.2, 2.3, 8.7, 4.5], ona treba da vrati broj 2.3.

	Function ManjiSusjedi(X() as Single) as Single
		Dim IND as Integer
		For I = UBound(X)-1 to LBound(X)+1
			If X(I) < X(I+1) And X(I) < X(I-1) Then
				ManjiSusjedi = X(I)
				IND = 1
				Exit For
			End If
		Next
		If IND = 0 Then
			ManjiSusjedi = X(UBound(X))
		End If
	EndFunction







-----------------------------------------------------------------------------------------
Dim X() as Integer
Dim X(0 to 10) as Integer

LBound(X) – pozicija prvog elementa
UBound(X) – pozicija posljednjeg elementa

For I = LBound(X) to Ubound(X)
X(I) – element niza X
-----------------------------------------------------------------------------------------


7.3 (4 poena) Napisati VBA funkciju ParniKvadrati koja za argument ima niz cijelih brojeva X i vraca zbir kvadrata svih parnih elemenata niza. 
	
Primjer: Ako se funkciji proslijedi niz [7,3,2,8,11], ona treba da vrati broj 68=22+82

	Function ParniKvadrati(X() as Integer) as Integer
		Dim S as Integer, I as Integer
		S = 0
		For I = LBound(X) to Ubound(X)
			If X(I) Mod 2 = 0 Then
				S = S+(X(I)^2)
			End If
		Next
		ParniKvadrati = S
	End Function

7.4 (4 poena) Napisati VBA funkciju Sredina koja za argument ima niz realnih brojeva X i koja vraća broj elemenata niza koji su veći od lijevog susjeda (prvi sa lijeva), a manji od desnog susjeda (prvi sa desna). Prvi i zadnji element ne uzimati u obzir.

Primjer: Ako se funkciji proslijedi niz [6.9, 1.1, 3.2, 4.4, 2.7, 4.5, 9.3], ona treba da vrati broj 2, jer elementi 3.2 i 4.5 zadovoljavaju traženi uslov.

	Function Sredina(X() as Double) as Integer
		Dim BR as Integer, I as Integer
		BR = 0
		For I = LBound(X)+1 to Ubound(X)-1
			If X(I) > X(I-1) And X(I) < X(I+1) Then
				BR = BR+1
			End If
		Next
		Sredina = BR
	End Function






7.5 (5 poena) Napisati VBA funkciju Balans koja za argumente ima niz cijelih brojeva X i cio broj K. Funkcija vraća True ako je niz uravnotežen u odnosu na broj K i False u suprotnom. Niz je uravnotežen u odnosu na broj K ako je broj elemenata niza manjih od K jednak broju elemenata većih od K.

Primjer: Niz X = [2,3,9,4,11,22] je uravnotežen u odnosu na broj K=7 jer postoje 3 elementa niza koji su veći od K i 3 koji su manji od K.

	Function Balans(X() as Integer, K as Integer) as Boolean
Dim BR1 as Integer, BR2 as Integer. I as Integer
BR1 = 0 : BR2 = 0
For I = LBound(X) to Ubound(X)
	If X(I) < K Then
		BR1 = BR1 + 1
	Else If X(I) > K Then
		BR2 = BR2 + 1
	End If
Next
If BR1 = BR2 Then
	Balans = True
Else
	Balans = False
End If
	End Function

7.6 (5 poena) Napisati VBA funkciju Pojavljivanje koja za argumente ima dva stringa P i Q. Funkcija treba da vrati koliko se karaktera stringa Q pojavljuje u stringu P. Primjer: 


Funkcija Pojavljivanje(“Abc7”,”d23x”) vraća 0, dok Pojavljivanje(“Abc7”,”bxb7x”) vraća broj 3 (pojavljuju se karakteri b, b i 7).


	Function Pojavljivanje(P As String, Q As String) As Integer
Dim I As Integer
		Dim J As Integer
		BR = 0
		For I = 1 To Len(Q)
			For J = 1 To Len(p)
				If Mid(Q,I,1) = Mid(P,J,1) Then
					BR = BR + 1
				End If
			Next
		Next
		Pojavljivanje = BR
	End Function
 












7.7 (5 poena) Napisati VBA funkciju DvaStringa koja za argumente ima dva stringa P i Q. Ako su stringovi iste dužine, funkcija treba da vrati string dobijen nadovezivanjem posljednja dva karaktera stringa Q na prva dva karaktera stringa P. U suprotnom, funkcija treba da vrati string P sa obrnutim redosljedom karaktera. 


Primjer: Funkcija DvaStringa(“Abc7”,”d23x”) vraća string “Ab3x”, dok DvaStringa(“Abc7”,”x”) vraća string “7cbA”.


	Function DvaStringa(P As String, Q As String) As String
		S = ""
		If Len(P) = Len(Q) Then
			S = Right(Q, 2) + Left(P, 2)
		Else
			For I = 1 To Len(P)
				S = Mid(P,I,1) + S
			Next
		End If
	End Function




7.8 (4 poena) Napisati VBA funkciju PrvaPojava koja za argumente ima string S i karakter K. Funkcija treba da vrati poziciju prve pojave karaktera K u stringu S. Ukoliko karakter K ne postoji u stringu S, funkcija treba da vrati -1. 


Primjer: Funkcija PrvaPojava("Abc7ef7g","7") vraća broj 4, PrvaPojava("Abc7ef7g","8") vraća -1.


	Function PrvaPojava(S as String, K as String) as String
		Dim I as Integer, IND as Integer
		
		For I = 1 To Len(S)
			If Not Mid(S,I,1) Like K Then
				PrvaPojava = -1
			Elseif Mid(S,I,1) Like K Then
				PrvaPojava = I
				Exit For
			End If
		Next
	End Function






















7.9 (4 poena) Napisati VBA funkciju Izraz1 koja za argument ima niz realnih brojeva dvostruke preciznosti X i koja vraća zbir kvadrata svih negativnih elemenata i kubova svih nenegativnih (nula i pozitivni) elemenata niza.


Primjer: Ako se funkciji proslijedi niz [7.1, -3.2, 2.5, -1, -2.7], ona treba da vrati vrijednost izraza 7.13 + (-3.2)2 + 2.53 + (-1)2+ (-2.7)2.


	Function Izraz1(X() as Double) as Double
		S as Double
		For I = LBound(X) to Ubound(X)
			If X(I) < 0 Then
				S = S + X(I)^2
			Else If X(I) => 0 Then
				S =  S + X(I)^3
			End If
		Next
		Izraz1 = S
	End Function




7.10 (4 poena) Napisati VBA funkciju ParniKvadrati koja za argument ima niz cijelih brojeva X i vraća zbir kvadrata svih parnih elemenata niza.


Primjer: Ako se funkciji proslijedi niz [7, 3, 2, 8, 11], ona treba da vrati broj 68 = 22+ 82.


	Function ParniKvadratui(X() as Integer) as Integer
		S as Integer
		S = 0
		For I = LBound(X) To Ubound(X)
			If X(I) Mod 2 = 0 Then
				S = S + (X(I)^2)
			End If
		Next
		ParniKvadrati = S
	End Function
	






















	
7.11 (4 poena) Napisati VBA funkciju PrviNegativan koja za argument ima niz Integer brojeva i koja vraća prvi negativan element na koji naiđe počinjući od prvog elementa. Ako nema negativnih elemenata, funkcija vraća prvi element niza.


Primjer: Ako se funkciji proslijedi niz [1,2,-2,3,4,-5,6], ona treba da vrati broj -2.


	Function PrviNegativan(niz() as Integer) as Integer
		For I = LBound(X) To UBound(X)
			If X(I) < 0 Then
				N = X(I)
				Exit For
			Else 
				N = X(LBound(X))
			End If
		Next
		PrviNegativan = N
	End Function




7.12 (4 poena) Napisati VBA funkciju PosljednjiParan koja za argument ima niz Integer brojeva i koja vraća posljednji paran element na koji naiđe. Ako nema parnih elemenata, funkcija vraća zadnji element niza.


Primjer: Ako se funkciji proslijedi niz [1,3,10,3,4,5,6,7], ona treba da vrati broj 6.


	Function PoslednjiParan(nix() as Integer) as Integer
		N as Integer
		For I=UBound(X) To LBound(X)
			If X(I) Mod 2 = 0 Then
				N = X(I)
				Exit For
			Else 
				N = X(UBound(X))
			End If
		Next
		PoslednjiParan = N
	End Function































For I = 1 to R.Rows.Count   					''' Redovi u opsegu R
			For J = 1 to R.Columns.Count		''' Kolone u opsegu R
				R.Cells(I,J).Value

N = ThisWorkbook.Worksheets.Count   	''' U slucaju da radi jedan radni list, stavljamo 
''' umjesto N broj, a ako je svakog radnog lista 
''' onda idemo sa For petljom
Set R = ThisWorkbook.Worksheets(N).Range(“A1:C10”)  

-----------------------------------------------------------------------------------------

8.1 (5 poena) Napisati VBA funkciju Dvanaest koja za svoj argument ima opseg ćelija R. Funkcija treba da vrati broj ćelija u opsegu R koje imaju tačno 12 karaktera.

	Function Dvanaest(R as Range) as Integer
		Dim I as Integer, J as Integer, BR as Integer
		For I = 1 to R.Rows.Count   			''' Redovi u opsegu R
			For J = 1 to R.Columns.Count		''' Kolone u opsegu R
				If Len(R.Cells(I,J).Value) = 12 Then
					BR = BR+1
				End If
			Next
		Next
		Dvanaest = BR
	End Function























8.2 (5 poena) Napisati VBA proceduru Opsezi koja za argumente ima dva opsega ćelija R1 i R2 i koja upisuje tekst iz prve ćelije opsega R1 (presjek prve vrste i prve kolone) u svaku praznu ćeliju opsega R2. U slučaju da je prva ćelija opsega R1 prazna, u sve prazne ćelije opsega R2 upisati Vaš broj indeksa.

	Sub Opsezi(R1 as Range, R2 as Range) 
		Dim I as Integer, J as Integer
		If R1.Cells(1,1).Value = “” Then
			For I = 1 to R2.Rows.Count   		
				For J = 1 to R2.Columns.Count
					If R2.Cells(I,J).Value Like “” Then	
						R2.Cells(I,J).Value = “96/18”
					End If
				Next
			Next
		Else
			For I = 1 to R2.Rows.Count   		
				For J = 1 to R2.Columns.Count	
					If R2.Cells(I,J).Value Like “” Then 
						R2.Cells(I,J).Value = R1.Cells(1,1).Value
					End If
				Next
			Next
		End If
	End Sub


8.3 (5 poena) Napisati VBA funkciju Opsezi koja za argumente ima dva opsega ćelija R1 i R2, koji imaju isti broj vrsta i kolona (ne provjeravati). Funkcija treba da poredi opsege R1 i R2 ćeliju po ćeliju (prvu ćeliju opsega R1 sa prvom ćelijom opsega R2, drugu ćeliju opsega R1 sa drugom ćelijom opsega R2 itd.) i da vrati koliko puta je sadržaj odgovarajućih ćelija isti.

	Function Opsezi(R1 as Range, R2 as Range) as Integer
		Dim I as Integer, J as Integer, BG as Integer

For I = 1 to R2.Rows.Count   		
			For J = 1 to R2.Columns.Count
				If R1.Cells(I,J).Value = R2.Cells(I,J).Value Then
					BR = BR + 1
				End If
			Next
		Next
		Opsezi = BR
	End Function




8.4 (5 poena) Napisati VBA funkciju VisePraznih koja za argumente ima dva opsega ćelija, R1 i R2. Funkcija vraća broj 1 ako opseg R1 ima više praznih ćelija od R2, broj -1 ako opseg R2 ima više praznih ćelija od R1, i 0 ako imaju isti broj praznih ćelija.

	Function VisePraznih(R1 as Range, R2 as Range) as Integer
		Dim I as Integer, J as Integer, BR1 as Integer, BR2 as Integer
		
	For I = 1 To R1.Rows.Count
		For J = 1 To R1.Columns.Count
				If R1.Cells(I,J).Value Like “” Then
					BR1 = BR1 + 1
				End If
			Next
		Next
	For I = 1 To R2.Rows.Count
		For J = 1 To R2.Columns.Count
				If R2.Cells(I,J).Value Like “” Then
					BR2 = BR2 + 1
				End If
			Next
		Next
		If BR1 > BR2 Then
			VisePraznih = 1
		Elseif BR1 < BR2 Then
			VisePraznih = -1
		Elseif BR1 = BR2 Then
			VisePraznih = 0
		End If
End Function





















8.5 (5 poena) Napisati VBA proceduru PraznoPrezime koja za argument ima opseg ćelija R i koja upisuje tekst iz prve ćelije opsega (presjek prve vrste i prve kolone) u svaku praznu ćeliju tog opsega. U slučaju da je prva ćelija opsega prazna, u sve prazne ćelije datog opsega upisati Vaše prezime.

Sub PraznoPrezime(R as Range)
	Dim I as Integer, J as Integer
	If R.Cells(1,1).Value Like “” Then
		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				If R.Cells(I,J).Value Like “” Then
					R.Cells(I,J).Value = “Marijanovic”
				End If
			Next
		Next
	Else
		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				If R.Cells(I,J).Value Like “” Then
					R.Cells(I,J).Value = R.Cells(1,1).Value
				End If
			Next
		Next
End Sub

8.6 (5 poena) Napisati Excel VBA funkciju Celije koja će proći kroz opseg A1:C10 svakog radnog lista ThisWorkbook radne sveske i provjeriti da li je u bar jednoj ćeliji upisan broj 505. U slučaju da postoji bar jedna takva ćelija, funkcija treba da vrati broj 1. U suprotnom, funkcija vraća broj 0.

	Function Celija() as Integer
		N = ThisWorkbook.Worksheets.Count                   ''' Broj radnih listova

		For k = 1 To N
		Set R = ThisWorkbook.Worksheets(k).Range(“A1:C10”)  ''' Setujemo opseg _
                                                                ''' u radnom listu
		For I = 1 to R.Rows.Count   		
			For J = 1 to R.Columns.Count
				If IsNumeric(R.Cells(I,J.Value) and R.Cells(I,J).Value = 505 Then
					BR = BR+1
				Endif
			Next
		Next
		Next
		If BR > 0 Then
			Celije = 1
		Else
			Celije = 0
		End If
	End Function

8.7 (5 poena) Aktivna radna sveska sadrži radni list Test. Napisati VBA proceduru UpisBroja koja za argument ima realan broj X. Procedura treba da upiše broj X u svaku praznu ćeliju opsega A1:E8 radnog lista Test aktivne sveske.

	Sub UpisBroja(X as Double)
		Dim J as Integer, I as Integer
		Set R = ActiveWorkBook.Worksheets("Test").Range("A1:E8)

		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				If R.Cells(I,J).Value Like “” Then
					R.Cells(I,J).Value = X
				End If
			Next
		Next
	End Sub

8.8 (5 poena) Napisati VBA proceduru PrvaDvaNadovezi koja prolazi kroz sve ćelije opsega B2:C10 prvog radnog lista aktivne sveske i formira string nadovezujući prva dva karaktera svake ćelije. Dobijeni string prikazati pomoću Message box-a.

	Sub PrvaDvaNadovezi()
		Dim J as Integer, I as Integer, S as String
		S = “”
		Set R = ActiveWorkBook.Worksheets(1).Range("B2:C10)

		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				S = S & Left(R.Cells(I,J).Value,2)
			Next
		Next
		MsgBox S
	End Sub

8.9 (5 poena) Napisati VBA proceduru NadoveziPosljednje koja prolazi kroz sve ćelije opsega C1:C10 prvog radnog lista aktivne sveske i formira string nadovezujući posljednje karaktere svake ćelije. Dobijeni string prikazati pomoću Message box-a.

Sub NadoveziPosljednje()
		Dim J as Integer, I as Integer, S as String
		S = “”
		Set R = ActiveWorkBook.Worksheets(1).Range("C1:C10”)

		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				S = S & Right(R.Cells(I,J).Value,1)
			Next
		Next
		MsgBox S
	End Sub

8.10 (5 poena) Napisati VBA proceduru PrvaDvaNadovezi koja prolazi kroz sve ćelije opsega B2:C10 prvog radnog lista aktivne sveske i formira string nadovezujući prva dva karaktera svake ćelije. Dobijeni string prikazati pomoću Message box-a.

	Sub PrvaDvaNadovezi()
		Dim J as Integer, I as Integer, S as String
		
		Set R = ActiveWorkBook.Worksheets(1).Range(“B2:C10”)
		
		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				S = S & Left(R.Cells(I,J).Value,2)
			Next
		Next
		MsgBox S
	End Sub

8.10 (5 poena) Napisati VBA proceduru PocinjeKaoPrva koja određuje koliko ima ćelija u opsegu A1:C10 prvog radnog lista aktivne sveske koje počinju istim karakterom kao i ćelija A1 tog opsega. Dobijeni broj prikazati u Immediate prozoru.

	Sub PocinjeKaoPrva()
		Celija = 0
		Dim I as Integer, J as Integer
	
		Set O = ActiveWorkBook.Worksheets(1).Range(“A1”)
		Set R = ActiveWorkBook.Worksheets(1).Range(“A1:C10”)

		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				If Left(R.Cells(I,J).Value,1) = Left(O.Cells(I,J).Value,1) Then
					BR = BR + 1
				End If
			Next
		Next
	End Sub
				
	











8.11 (5 poena) Napisati VBA proceduru Nadovezi koja za argument ima string STR i koja taj string nadovezuje na svaku ćeliju opsega A2:G10 drugog radnog lista aktivne sveske koja ima manje od 15 karaktera.
	Sub()
		Dim I as Integer, J as Integer, STR as String
		Set R = ActiveWorkBook.Worksheets(2).Range(“A2:G10”)
		For I = 1 To R.Rows.Count
			For J = 1 To R.Columns.Count
				If Len(R.CelLs(I,J).Value) < 15 Then
					R.Cells(I,J).Value = STR & R.Cells(I,J).Value
				End If
			Next
		Next
	End Sub

	

8.12 (6 poena) Napisati Excel VBA funkciju SviKarakteri koja za argument ima opseg ćelija R. Funkcija treba da odredi i vrati ukupan broj karaktera u svim ćelijama tog opsega. U slučaju da su sve ćelije prazne, funkcija treba da u prvu ćeliju tog opsega upiše tekst "Sve prazno".

	Function SviKatakteri(R as Range)
		Dim I as Integer, J as Integer, S 
		For I = 1 To R.Rows.Count
			For J = 1 To R.Coluns.Count
				If R.Cells(I,J).Value Like “” Then
					R.Cells(I,J).Value = “Sve Prazno”
				Else
					S = S + Len(R.Cells(I,J).Value)
			End If
		Next
		SviKarakteri = S
	End Function
