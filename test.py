import page2ppt
if __name__ == "__main__":
    url = "http://www.ccview.net/htm/xiandai/zzq/zzqsw002.htm"
    ppt = page2ppt.page2ppt(url, ppt_pages=5)
    ppt.convert2ppt()
    print "done"
