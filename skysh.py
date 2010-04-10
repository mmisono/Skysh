#!/usr/bin/python
# -*- coding: utf-8 -*-

import Skype4Py
from termcolor import *
import webbrowser
import sys
import os
import cmd
import commands
import re
import readline
import threading


class Skysh(cmd.Cmd):

  def __init__(self):
    cmd.Cmd.__init__(self)
    self.intro  = "Skysh version 0.0.1"
    self.prompt = Prompt()
    self.skype = Skype4Py.Skype()
    if not self.skype.Client.IsRunning:
      print 'Starting Skype...'
      self.skype.Client.Start()
      while not self.skype.Client.IsRunning:
        pass
    self.skype.OnMessageStatus = self.OnMessageStatus
    self.skype.OnAsyncSearchUsersFinished = self.OnAsyncSearchUsersFinished
    self.skype.Attach()
    self.urllist = []
    self.chat = ''

  
  def do_help(self,ignore):
    print """
    COMMAND \t\t\t DIRECTIONS
    add <SkypeID> \t\t Add user to current chat
    bookmark \t\t\t Bookmark current chat
    bookmarks \t\t\t Show List of bookmarks and change current chat
    chat <SkypeID> \t\t Begin chat with <SkypeID>
    exit \t\t\t exit 
    friends \t\t\t Show friends
    help \t\t\t Show this message
    kick <SkypeID> \t\t Kick user from current chat
    ls \t\t\t\t Show list of active chats
    members \t\t\t Show members of current chat and show status
    mood [Message] \t\t Change mood message or show current mood message
    open \t\t\t Show lists of recent url which is in message and open url
    profile <SkypeID> \t\t\t Show profile
    recent \t\t\t Show lists of recent chat and change current chat
    search <username> \t\t\t Search user
    status \t\t\t Show current status and change 
    unbookmark \t\t\t Unbookmark current chat
    """


  #Send message
  def default(self, line):     
    if not self.chat:
      self.error("Please chat someone")
    else:
      self.chat.SendMessage( line.decode('utf-8') )


  def emptyline(self):
    pass


  #Add user to current chat
  def do_add(self,username):
    try:
      user = self.skype.User(Username=username)
      self.chat.AddMembers(user)
    except Skype4Py.errors.SkypeError:
      pass


  #Bookmark Current chat
  def do_bookmark(self,ignore):
    try:
      if self.chat:
        self.chat.Bookmark()
    except Skype4Py.errors.SkypeError:
      pass


  #Show Bookmarks
  def do_bookmarks(self,ignore):
    bookmarks = self.skype.BookmarkedChats
    for i in xrange(bookmarks.Count):
      print "%2d: %s" % (i,bookmarks.Item(i).FriendlyName)

    if bookmarks.Count > 0:
      num = self.select(bookmarks.Count)
      if num >= 0:
        self.chat = self.skype.BookmarkedChats.Item(num)

  

  #Create chat
  def do_chat(self,user):
    self.chat = self.skype.CreateChatWith(user)


  #Exit
  def do_exit(self,ignore):
    sys.exit(1)

  def do_EOF(self,ignore):
    print '\n',
    sys.exit(1)


  #Show friends
  def do_friends(self,ignore):
    cnt = 0
    for user in self.skype.Friends:
      print "%2d: %s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
      cnt += 1


  #Kick user form current chat
  def do_kick(self,username):
    try:
      self.chat.Kick(username)
    except Skype4Py.errors.SkypeError:
      pass


  #Show list of active Chat
  def do_ls(self,arg):
    chat = self.skype.ActiveChats
    try:
      for i in xrange(chat.Count):
        print "%2d: %s" % (i,chat.Item(i).FriendlyName)
    except Skype4Py.errors.SkypeError:
      pass

    if chat.Count > 0:
      num = self.select(chat.Count)
      if num >= 0:
        self.chat = self.skype.ActiveChats(num)

  #Print Current Chat Member
  def do_members(self,ignore):
    if self.chat:
      cnt = 0
      for i in xrange(self.chat.Members.Count):
        user = self.chat.Members.Item(i)
        print "%2d: %s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
        cnt += 1


  #Show or Chanege mood text:
  def do_mood(self,mes):
    if not mes:
      print self.skype.CurrentUserProfile.MoodText
    else:
      self.skype.CurrentUserProfile.MoodText = mes.decode('utf-8')


  #Show recent url which is in messages and open url
  def do_open(self,ignore):
    cnt = 0
    if len(self.urllist) == 0:
      return
    for url in self.urllist:
      print "%2d: %s" % (cnt,url)
      cnt += 1

    if cnt > 0:
      num = self.select(len(self.urllist))
      if num >= 0:
        webbrowser.open(self.urllist[num])


  #Show user profile
  def do_profile(self,username):
    user = self.skype.User(username)
    print "%11s: %s" % ('FullName',user.FullName      )
    print "%11s: %s" % ('Handle',user.Handle          )
    print "%11s: %s" % ('DisplayName',user.DisplayName)
    print "%11s: %s" % ('MoodText',user.MoodText      )
    print "%11s: %s" % ('About',user.About            )
    print "%11s: %s" % ('Language',user.Language      )
    print "%11s: %s" % ('Birthday',user.Birthday      )
    print "%11s: %s" % ('Country',user.Country        )
    print "%11s: %s" % ('City',user.City              )
    print "%11s: %s" % ('Province',user.Province      )
    print "%11s: %s" % ('PhoneHome',user.PhoneHome    )
    print "%11s: %s" % ('PhoneMobile',user.PhoneMobile)
    print "%11s: %s" % ('PhoneOffice',user.PhoneOffice)
    print "%11s: %s" % ('Homepage',user.Homepage      )
    print "%11s: %s" % ('Sex',user.Sex                )


  #Show list of recent chat and change
  def do_recent(self,ignore):
    chats = self.skype.RecentChats
    for i in xrange(chats.Count):
      print "%2d: %s" % (i,chats.Item(i).FriendlyName)
      
    if chats.Count > 0:
      num = self.select(chats.Count)
      if num >= 0:
        self.chat = self.skype.RecentChats.Item(num)

  #Search user
  def do_search(self,username):
    self.skype.AsyncSearchUsers(username)


  #Show and Change user status
  def do_status(self,ignore):
    status = ['ONLINE','SKYPEME','AWAY','NA','DND','INVISIBLE','OFFLINE']
    status_ja = ['オンライン','SkypeMe!','一時退席中','退席中','取り込み中','ログイン状態を隠す','オフライン']
    print 'Now status: ' + self.skype.CurrentUserStatus
    for i in xrange(len(status)):
      print "%2d: %s(%s)" % (i,status[i],status_ja[i])

    if i > 0:
      num = self.select(len(status))
      if num >= 0:
        self.skype.ChangeUserStatus(status[num])


  #UnBookmark Current chat
  def do_unbookmark(self,ignore):
    try:
      if self.chat:
        self.chat.Unbookmark()
    except Skype4Py.errors.SkypeError:
      pass


  #Skype Events
  def OnMessageStatus(self,mes, status):
    if status == 'RECEIVED':
      threading.Thread(target=self.printMessage(mes)).start()

  def OnAsyncSearchUsersFinished(self,cookie,users):
    cnt = 0
    print '\r' + ' ' * 100 + '\r',
    for user in users:
      print "%2d:%s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
      cnt += 1
    if self.chat:
      print '(' + skype.chat.FriendlyName.encode('utf-8') + ')',
    print self.skype.CurrentUserProfile.FullName +  ':',readline.get_line_buffer(),
    sys.stdout.flush()
  

  #Print Error Message
  def error(self,mes):
    print colored("[ERROR] " + mes,'red')


  def select(self,max):
    try:
      num = int(raw_input('Enter number: '))
    except ValueError:
      return -1
    if num < 0 or num >= max:
      self.error('Invalid Number')
      return -1
    else:
      return num


  #Select color
  def selectColor(self,chat,handle):
    color = ['green','magenta','cyan','blue','yellow']
    if self.chat:
      for i in xrange(self.chat.Members.Count):
        if chat.Members.Item(i).Handle == handle:
          return color[i % len(color)]
    else:
      return color[0]


  def appendURL(self,url):
    if len(self.urllist) > 10:
      self.urllist = self.urllist[len(self.urllist)-10:]
    self.urllist.append(url)


  def printMessage(self,mes):
    print '\r' * int(commands.getoutput('echo $COLUMNS')), #Clear current line
    color = self.selectColor(mes.Chat,mes.FromHandle)
    self.getURL(mes.Body)
    if self.chat == mes.Chat:
      print '\r(' + mes.Chat.FriendlyName + ') ' + colored(  mes.FromDisplayName + ': '  +  mes.Body ,color)
    else:
      print colored('\r(' + mes.Chat.FriendlyName + ') ' + mes.FromDisplayName + ': ' +  mes.Body ,color)
    if self.chat:
      print '\r(' + self.chat.FriendlyName.encode('utf-8') + ')',
    print self.skype.CurrentUserProfile.FullName.encode('utf-8') +  ': ' + readline.get_line_buffer(),
    sys.stdout.flush()
  
  
  #if url is found,append list
  def getURL(self,mes):
    urlpattern = [re.compile("([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}|(((news|telnet|nttp|file|http|ftp|https)://)|(www|ftp)[-A-Za-z0-9]*\\.)[-A-Za-z0-9\\.]+)(:[0-9]*)?"), 
                  re.compile("([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}|(((news|telnet|nttp|file|http|ftp|https)://)|(www|ftp)[-A-Za-z0-9]*\\.)[-A-Za-z0-9\\.]+)(:[0-9]*)?/[-A-Za-z0-9_\\$\\.\\+\\!\\*\\(\\),;:@&=\\?/~\\#\\%]*[^]'\\.}>\\),\\\"]")
                  ]

    for p in urlpattern:
      urls = p.findall(mes)
      if urls:
        for url in urls:
          self.appendURL(url[0])



class Prompt():
  def __repr__(self):
    string = "\r"
    if skype.chat:
      string += '(' + skype.chat.FriendlyName.encode('utf-8') + ') '
    string += skype.skype.CurrentUserProfile.FullName.encode('utf-8') + ": "
    return string



if __name__ == '__main__':
  skype = Skysh()
  threading.Thread(target=skype.cmdloop()).start()
