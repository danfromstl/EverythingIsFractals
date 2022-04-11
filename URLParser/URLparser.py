import csv
import re

with open('DarknetDiaries-CVS_Export.csv', encoding='Latin1', newline='') as csvfile:
    episodereader = csv.reader(csvfile, delimiter=',', quotechar='"')
    episode_count = 0
    globalurl_count = 0
    twitterhandle_count = 0

    urlsearch = re.compile("(?P<url>[^\s]+\.[org|com|tv][\S]+)")

    # allPersons = []
    for row in episodereader:
        # conditional that ingores header row
        if(row[0] == "Number"):
            pass
        else:
            ep_title = row[1]
            ep_desc = row[6]

            currentURLs = re.findall(urlsearch, ep_desc)
            url_count = len(currentURLs)
            
            if(url_count > 0):
                # print(f'Title:{ep_title}\n')
                # print(f'Description: {ep_desc}\n')
                print(f'Found {url_count} urls in {ep_title}')
                # print(f'URLs: {currentURLs}\n\n')
                for epURL in currentURLs:
                    result = re.match(".+twitter.+",epURL)
                    #print(result)
                    if(result!=None):
                        print(epURL)
                        twitterhandle_count += 1
                print('')
                
            episode_count += 1
            globalurl_count += url_count
            
#           if(row[10] not in firstOrg.leaderStringIDlist and row[10] != ""):
#               firstOrg.leaderStringIDlist.append(row[10])            
#           
#           firstOrg.peopleTree.append(personStruct(row[0],row[4],row[6],row[8],row[9],row[10]))
#           episode_count += 1

    
    print(f'\n------\nEpisode list - first pass complete\nRead {episode_count} episodes')
    print(f'Found {globalurl_count} URLs')
    print(f'Found {twitterhandle_count} Twitter handles')