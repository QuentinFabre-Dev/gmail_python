
def getInfos(ctx,feuille,l):
    i = 0
    payload = []
    for part in ctx.walk():
        payload.append(part.get_payload())
        i = i + 1

    pyl = payload[2].replace("=0D","").replace("=0D","").replace("=C3=B4","o").replace("-","")
    # payload[1] si mail transferer
    idxPas = pyl.find('Passenger')
    idxTel = pyl.find('Mobile')
    idxBeSure = pyl.find('Be sure')
    idxDate = pyl.find('following ride')
    idxBookingNumber = pyl.find('Booking number')
    idxFrom = pyl.find('From')
    idxTo = pyl.find('To')
    idxDistance = pyl.find('Distance')
    


    passenger = pyl[idxPas+10:idxTel-2]
    mobile = pyl[idxTel+7:idxBeSure]
    laDate = pyl[idxDate+25:idxBookingNumber]
    depart = pyl[idxFrom+5:idxTo]
    arrive = pyl[idxTo+3:idxDistance]

    # date = pyl[idxDate+30:747].replace(">","")
  
    print("=================================================================================================================================================")
    print("Passenger : "+passenger)
    print("Mobile : "+mobile)
    print("Date : "+laDate)
    print("From : "+depart)
    print("To : "+arrive)

    print("["+str(l)+"]")
    print("=================================================================================================================================================")

    feuille.write(l, 0, passenger)
    feuille.write(l, 1, mobile)
    feuille.write(l, 2, depart)
    feuille.write(l, 3, arrive)
    feuille.write(l, 4, laDate)