w1 = ["fin","reconcile","throttle","spectator","jihad","umpire","uprising","steroid","crumb","edible","provincial","puddle","gravel","vengeance","martyr","tug","disciple","integral","sneaky","drape","imaging","sergeant","cram","erode","jab","buff","catalyst","oppression","turbine","artillery","bilateral","gruesome","prospective","ripple","rout","eject","propietary","sage","topple","tribunal","populist","flex","dangle","arsenal","exploitation","conjure","mould"]
w2 = ["existential","parlour","sizzle","winthdrawn","marvel","sitcom","puncture","interception","meticulous","thrash","tremble","cradle","curd","assertion","perch","fallout","gravy","residue","procession","slew","caste","imitate","engulf","taunt","contention","parasite","friction","hush","ravage","barricade","rigorous","dissent","lithium","bleak","hoof","mound","abolish","neuron","ooze","choreography","diversion","patriotic","scorch","mist","spat","vigorous","congregation"]
w3 = ["thwart","decor","upbringing","nudge","bulter","hoax","optical","unearth","inhibit","blatant","glide","shag","glitch","pundit","immerse","scalp","zebra","avenge","reconciliation","amulet","genetically","fascist","submerge","envoy","crypto","freight","craze","hygiene","succumb","swoop","bullish","shealth","perpetual","relish","reptile","shack","cyclone","oppress","ridicule","syndicate","brunette","onward","kin","shiver","testament","censor","dunk"]
w4 = ["thorn","proactive","overrun","cloak","compassionate","spook","edgy","tsunami","exterminate","mow","saturate","ale","sidekick","loophole","sling","hologram","repeal","rumble","wreak","heed","anal","tangible","hunk","communal","outward","confinement","composite","obscene","relegate","slime","antidote","imperial","apprentice","bladder","fishy","pixel","solicitor","precinct","torque","wretch","sleigh","exacerbate","nip","preside","deflect","plight","taint"]

def scrape_example(url, word):
    opener = urllib.request.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    response = opener.open(url + word)
    html_contents = response.read()
    soup = BeautifulSoup(html_contents, 'html.parser')
    if len(soup.find_all("span", class_="ipa dipa lpr-2 lpl-1"))>0:
        pronouciation = soup.find_all("span", class_="ipa dipa lpr-2 lpl-1")[0].get_text().strip()
    else:
        pronouciation = 'none'
    return pronouciation

lines = ["word,pron,word,pron,word,pron,word,pron"]    
for i in range(len(w1)):
    p1 = w1[i]+','+ scrape_example(url, w1[i])
    p2 = w2[i]+','+ scrape_example(url, w2[i])
    p3 = w3[i]+','+ scrape_example(url, w3[i])
    p4 = w4[i]+','+ scrape_example(url, w4[i])
    lines.append(','.join([p1,p2,p3,p4]))

with open('/Users/yun/test1.csv', 'w') as f:
    f.write('\n'.join(lines))


import xlrd
workbook = xlrd.open_workbook("~/Downloads/pronunciation/voc19000-20000.xlsx")
sheet_names = workbook.sheet_names()
group = sheet_names[0]
worksheet = workbook.sheet_by_name(group)
word_list = [[]] * (int(worksheet.ncols/2)+1)
pron_list = [[]] * (int(worksheet.ncols/2)+1)
url_list = ['https://dictionary.cambridge.org/us/dictionary/english/', 'https://www.ldoceonline.com/dictionary/']
url = url_list[0]
for i in range(0, worksheet.ncols, 2):
    col = worksheet.col(i)
    print(i)
    j = int(i/2)
    word_list[j] = [item.value.strip().replace(" ", "") for item in col if isinstance(item.value, str) and item.value.strip().replace(" ", "")]
    pron_list[j] = [scrape_example(url, w) if w.isascii() else 'none' for w in word_list[j]]

lines_list = []
for i in range(len(word_list[0])):
    line = []
    for j in range(len(word_list)):
        if i < len(word_list[j]):
            line.append(word_list[j][i])
            line.append(pron_list[j][i])
    lines_list.append(','.join(line))
with open('/Users/yun/8.csv', 'w') as f:
    f.write('\n'.join(lines_list))       


