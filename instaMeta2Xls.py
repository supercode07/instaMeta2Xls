import os
import json
import xlwt
import lzma

def dumpOneJson2Xml(xlsheet, load_dict, num):
    No = num
    postTime = load_dict['node']['taken_at_timestamp']
    Username = load_dict['node']['owner']['username']
    FollowersCount = load_dict['node']['owner']['edge_followed_by']['count']
    FollowingCount = load_dict['node']['owner']['edge_follow']['count']
    FollowerFollowingRate = FollowersCount / FollowingCount
    LikesCount = load_dict['node']['edge_liked_by']['count']
    Likerate = LikesCount/FollowersCount
    CommentCount = load_dict['node']['edge_media_to_comment']['count']
    PostType = load_dict['node']['__typename']
    Caption = load_dict['node']['edge_media_to_caption']['edges']

    #dump to excel file
    xlsheet.write(num, 0, No)
    xlsheet.write(num, 1, postTime)
    xlsheet.write(num, 2, Username)
    xlsheet.write(num, 3, FollowersCount)
    xlsheet.write(num, 4, FollowingCount)
    
    formula = "D" + str(num+1) + "/" + "E" + str(num+1)
    style_percent = xlwt.XFStyle()
    style_percent.num_format_str = '0.00%' 
    xlsheet.write(num, 5, xlwt.Formula(formula), style_percent)
    
    xlsheet.write(num, 6, LikesCount)
    
    formula = "G" + str(num+1) + "/" + "D" + str(num+1)
    xlsheet.write(num, 7, xlwt.Formula(formula), style_percent)
    
    xlsheet.write(num, 8, CommentCount)
    xlsheet.write(num, 9, PostType)
    xlsheet.write(num, 10, Caption)


#create a xml file
xlfile = xlwt.Workbook()
xlsheet = xlfile.add_sheet('postmetadata')
row0 = ['No', 'Post Date/Time', 'Username', 'Followers Count', 'Following Count', 'Follower/Following', 'Like Count', 'Like rate', 'Comment Count', 'Post Type', 'Caption']

style = xlwt.XFStyle()

font = xlwt.Font()
font.bold = True
font.height = 12 * 20

borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN

pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']


style.font = font
style.borders = borders
style.pattern = pattern

for i in range(0, len(row0)):
    xlsheet.write(0, i, row0[i], style)


postnum = 1
js_list = os.listdir('.\\')
for jsonFileName in js_list:
    if os.path.splitext(jsonFileName)[1] == '.xz':
        raw_file = lzma.open(jsonFileName)
        jsonFile = raw_file.read()
        load_dict = json.loads(jsonFile)
        dumpOneJson2Xml(xlsheet, load_dict, postnum)
        postnum = postnum+1

xlfile.save('metadata.xls')

