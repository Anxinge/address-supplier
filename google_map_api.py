import pgeocode

def find_code(postcode1):
    dist = pgeocode.GeoDistance('AU')
    #postcode1 = '4208'
    postcode2 = ['4209','3350','3355','3550','3175']
    distance_list = []
    for tempcode in postcode2:
        distance=dist.query_postal_code(postcode1, tempcode)
        print(distance)
        distance_list.append(distance)
    m = min(distance_list)
    
    i = 0
    for temp in postcode2:
        if dist.query_postal_code(postcode1,temp) == m :
            n = i
            resultcode = temp
            print(i,temp)
        i += 1
    return n,resultcode

n,r = find_code('2700')