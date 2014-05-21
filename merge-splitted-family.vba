Sub merge_splitted_family()

    Dim previous As String
    Dim comp_a As String
    Dim comp_b As String
    Dim cpt As Integer
    Dim deletion As Boolean
    Dim debugggy As String
        
    previous = "Noth92jlD298D1D1298DJhdIHDSIUHSKJHKJDHIU23H2HJKSNCXBCNXBBN"
       
    Dim a As Range, b As Range
    Set a = Range("D2", "E33")
    
    cpt = 2
    
    For Each b In a.Rows
                
        comp_a = previous
        comp_b = Range(Cells(b.Row, 4).Address) + Range(Cells(b.Row, 5).Address)
        
        If StrComp(comp_a, comp_b, vbBinaryCompare) = 0 And comp_b <> "" And comp_b <> "0" And comp_a <> "" And comp_a <> "0" Then
            
            Debug.Print " "
            Debug.Print "********** MATCH FOUND **********"
            Debug.Print comp_b + " - les chaines de comparaison sont égales"
            Debug.Print comp_b + " - chaine a : " + comp_a + " - chaine b : " + comp_b
            
            ' ECRIRE DATA THIS.RESP+THIS.CONJ -> RESP2+CONJ2 '
            Debug.Print comp_b + " - ligne 2 - Décalage des données en RESP2+CONJ2"
            
            Range(Cells(b.Row, 38).Address) = Range(Cells(b.Row, 1).Address) 'Lien parenté2'
            Range(Cells(b.Row, 39).Address) = Range(Cells(b.Row, 14).Address) 'Nom resp2'
            Range(Cells(b.Row, 40).Address) = Range(Cells(b.Row, 15).Address) 'Prenom resp2'
            Range(Cells(b.Row, 41).Address) = Range(Cells(b.Row, 17).Address) 'Telportable resp2'
            Range(Cells(b.Row, 42).Address) = Range(Cells(b.Row, 18).Address) 'Tel bureau resp2'
            Range(Cells(b.Row, 43).Address) = Range(Cells(b.Row, 19).Address) 'Email perso resp2'
            Range(Cells(b.Row, 44).Address) = Range(Cells(b.Row, 21).Address) 'Profession resp2'
            Range(Cells(b.Row, 45).Address) = Range(Cells(b.Row, 22).Address) 'Societe resp2'
            
            Range(Cells(b.Row, 46).Address) = Range(Cells(b.Row, 23).Address) 'Nom conjoint2'
            Range(Cells(b.Row, 47).Address) = Range(Cells(b.Row, 24).Address) 'Prenom conjoint2'
            Range(Cells(b.Row, 48).Address) = Range(Cells(b.Row, 25).Address) 'Telportable conj2'
            Range(Cells(b.Row, 49).Address) = Range(Cells(b.Row, 26).Address) 'Tel bureau conj2'
            Range(Cells(b.Row, 50).Address) = Range(Cells(b.Row, 27).Address) 'Email perso conj2'
            Range(Cells(b.Row, 51).Address) = Range(Cells(b.Row, 29).Address) 'Profession conj2'
            Range(Cells(b.Row, 52).Address) = Range(Cells(b.Row, 30).Address) 'Societe conj2'
            
            Range(Cells(b.Row, 53).Address) = Range(Cells(b.Row, 31).Address) 'Adresse 123 famille2'
            Range(Cells(b.Row, 54).Address) = Range(Cells(b.Row, 35).Address) 'Tel domicile 2'
            
            Range(Cells(b.Row, 55).Address) = Range(Cells(b.Row, 36).Address) 'CP-Ville'
            
            ' ECRIRE DATA CACHE -> RESP+CONJ '
            Debug.Print comp_b + " - ligne 2 - Ecriture des data backed en RESP1+CONJ1"
            
            Range(Cells(b.Row, 1).Address) = p_lienparente
            Range(Cells(b.Row, 14).Address) = p_nom_resp
            Range(Cells(b.Row, 15).Address) = p_prenom_resp
            Range(Cells(b.Row, 16).Address) = p_deuxifdifferent
            Range(Cells(b.Row, 17).Address) = p_telport_resp
            Range(Cells(b.Row, 18).Address) = p_telbureau_resp
            Range(Cells(b.Row, 19).Address) = p_emailperso_resp
            Range(Cells(b.Row, 21).Address) = p_profession_resp
            Range(Cells(b.Row, 22).Address) = p_sociecte_resp
        
            Range(Cells(b.Row, 23).Address) = p_nom_conj
            Range(Cells(b.Row, 24).Address) = p_prenom_conj
            Range(Cells(b.Row, 25).Address) = p_telport_conj
            Range(Cells(b.Row, 26).Address) = p_telbureau_conj
            Range(Cells(b.Row, 27).Address) = p_emailperso_conj
            Range(Cells(b.Row, 29).Address) = p_profession_conj
            Range(Cells(b.Row, 30).Address) = p_societe_conj
        
            Range(Cells(b.Row, 31).Address) = p_adresse123_famille
            Range(Cells(b.Row, 32).Address) = p_teldom
        
            Range(Cells(b.Row, 33).Address) = p_ville
            
            Debug.Print comp_b + " - Suppression de la premiere ligne de donnée"
            
            cpt = cpt - 1
            Range(Cells(cpt, 1), Cells(cpt, 60)).Delete
            cpt = cpt + 1
            
            Debug.Print " "
        End If
        
        ' COPIE CACHE DE RESP1 POUR NEXT TIME
        
        p_lienparente = Range(Cells(b.Row, 1).Address)
        p_nom_resp = Range(Cells(b.Row, 14).Address)
        p_prenom_resp = Range(Cells(b.Row, 15).Address)
        p_telport_resp = Range(Cells(b.Row, 17).Address)
        p_telbureau_resp = Range(Cells(b.Row, 18).Address)
        p_emailperso_resp = Range(Cells(b.Row, 19).Address)
        p_profession_resp = Range(Cells(b.Row, 21).Address)
        p_sociecte_resp = Range(Cells(b.Row, 22).Address)
        
        p_nom_conj = Range(Cells(b.Row, 23).Address)
        p_prenom_conj = Range(Cells(b.Row, 24).Address)
        p_telport_conj = Range(Cells(b.Row, 25).Address)
        p_telbureau_conj = Range(Cells(b.Row, 26).Address)
        p_emailperso_conj = Range(Cells(b.Row, 27).Address)
        p_profession_conj = Range(Cells(b.Row, 29).Address)
        p_societe_conj = Range(Cells(b.Row, 30).Address)
        
        p_adresse123_famille = Range(Cells(b.Row, 31).Address)
        p_teldom = Range(Cells(b.Row, 32).Address)
        
        p_ville = Range(Cells(b.Row, 33).Address)
        
        ' STOCKER EN CACHE POUR PROCHAINE ITERATION
        previous = comp_b
        
        comp_a = ""
        comp_b = ""
        
        cpt = cpt + 1
        
    Next
    
End Sub
