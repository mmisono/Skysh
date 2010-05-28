#!/usr/bin/python
# -*- coding: utf-8 -*-

#Author: mfumi <m.fumi760@gmail.com>
#Version: 0.0.3
#License: NEW BSD LICENSE
#  Copyright (c) 2010, mfumi
#  All rights reserved.
#
#  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#
#    * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#    * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the
#      documentation and/or other materials provided with the distribution.
#    * Neither the name of the tyru nor the names of its contributors may be used to endorse or promote products derived from this software without
#      specific prior written permission.
#
#  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
#  TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
#  CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
#  PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
#  LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
#  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#


import Skype4Py
from termcolor import *
import webbrowser
import sys
import os
import cmd
import re
import readline
import threading


class Skysh(cmd.Cmd):

  def __init__(self):
    cmd.Cmd.__init__(self)
    self.intro  = "Skysh version 0.0.3"
    self.prompt = Prompt()
    self.skype = Skype4Py.Skype()
    if not self.skype.Client.IsRunning:
      print 'Starting Skype...'
      self.skype.Client.Start()
      while not self.skype.Client.IsRunning:
        pass
    self.skype.Attach()
    self.skype.OnMessageStatus = self.OnMessageStatus
    self.skype.OnAsyncSearchUsersFinished = self.OnAsyncSearchUsersFinished
    self.skype.OnChatMembersChanged = self.OnChatMembersChanged
    self.skype.OnFileTransferStatusChanged = self.OnFileTransferStatusChanged
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
    file \t\t\t Show recent recieved files
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
      if line[0] == '\\':
        line = line[1:]
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
      print "\r%2d: %s" % (i,bookmarks.Item(i).FriendlyName)

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


  #Show Recent recieved files
  def do_file(self,ignore):
    i = 0
    cnt = 0
    fileno = []
    files = self.skype.FileTransfers
    while i < files.Count and cnt < 5:
      if files[i].Type == Skype4Py.fileTransferTypeIncoming:
        print "\r%2d %s: %s\t\tfrom %s" % (cnt,files[i].FinishDatetime.strftime("%Y-%m-%d %H:%M:%S"),files[i].FilePath.decode('utf_8'),files[i].PartnerDisplayName)
        fileno.append(i)
        cnt += 1
      i += 1
    if sys.platform == 'darwin' or sys.platform == 'win32':
      num = self.select(cnt)
      if num >= 0:
        if sys.platform == 'win32' :
          os.startfile(files[fileno[num]].FilePath)
        else :
          os.system("open %s" % files[fileno[num]].FilePath)


  #Show friends
  def do_friends(self,ignore):
    cnt = 0
    for user in self.skype.Friends:
      print "\r%2d: %s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
      cnt += 1


  #Kick user form current chat
  def do_kick(self,username):
    try:
      self.chat.Kick(username)
    except Skype4Py.errors.SkypeError:
      pass


  #Show list of active Chat
  def do_ls(self,arg):
    try:
      if self.skype.ActiveChats.Count == 0: # This case throw exception
        return
    except Skype4Py.errors.SkypeError:
      return

    #ActiveChats[0].Name contains all active chat name  ...why?
    chats = self.skype.ActiveChats[0].Name.split(' ')
    cnt = -1 
    for chat in chats:
      try:
        cnt += 1
        print "%2d:" % cnt,
#        print "%s" % self.skype.Chat(chat).FriendlyName
        print "%s" % Skype4Py.chat.Chat(self.skype,chat).FriendlyName
      except Skype4Py.errors.SkypeError:
        self.error("Can't get chat name")

    num = self.select(len(chats))
    if num >= 0:
      try:
#        self.chat = self.skype.Chat(chats[num])
        self.chat = Skype4Py.chat.Chat(self.skype,chats[num])
      except Skype4Py.errors.SkypeError:
        self.error("Can't change current chat")


  #Print Current Chat Member
  def do_members(self,ignore):
    if self.chat:
      cnt = 0
      for i in xrange(self.chat.Members.Count):
        user = self.chat.Members.Item(i)
        print "\r%2d: %s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
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
      print "\r%2d: %s" % (cnt,url)
      cnt += 1

    if cnt > 0:
      num = self.select(len(self.urllist))
      if num >= 0:
        webbrowser.open(self.urllist[num])


  #Show user profile
  def do_profile(self,username):
    user = self.skype.User(username)
    print "\r%11s: %s" % ('FullName',user.FullName      )
    print "\r%11s: %s" % ('Handle',user.Handle          )
    print "\r%11s: %s" % ('DisplayName',user.DisplayName)
    print "\r%11s: %s" % ('MoodText',user.MoodText      )
    print "\r%11s: %s" % ('About',user.About            )
    print "\r%11s: %s" % ('Language',user.Language      )
    print "\r%11s: %s" % ('Birthday',user.Birthday      )
    print "\r%11s: %s" % ('Country',user.Country        )
    print "\r%11s: %s" % ('City',user.City              )
    print "\r%11s: %s" % ('Province',user.Province      )
    print "\r%11s: %s" % ('PhoneHome',user.PhoneHome    )
    print "\r%11s: %s" % ('PhoneMobile',user.PhoneMobile)
    print "\r%11s: %s" % ('PhoneOffice',user.PhoneOffice)
    print "\r%11s: %s" % ('Homepage',user.Homepage      )
    print "\r%11s: %s" % ('Sex',user.Sex                )


  #Show list of recent chat and change
  def do_recent(self,ignore):
    chats = self.skype.RecentChats
    for i in xrange(chats.Count):
      print "\r%2d: %s" % (i,chats.Item(i).FriendlyName)
      
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
    print '\rNow status: ' + self.skype.CurrentUserStatus
    for i in xrange(len(status)):
      print "\r%2d: %s(%s)" % (i,status[i],status_ja[i])

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
    if status == 'RECEIVED' and len(mes.Body) > 0:
      threading.Thread(target=self.printMessage(mes)).start()

  def OnAsyncSearchUsersFinished(self,cookie,users):
    cnt = 0
    print "\x1b[2K",  #Clear current line
    for user in users:
      print "\r%2d:%s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
      cnt += 1
    self.printPrompt()
    print readline.get_line_buffer(),
    sys.stdout.flush()

  def OnChatMembersChanged(self,chat,mambers):
    cnt = 0
    print "\x1b[2K",  #Clear current line
    for member in members:
      print "\r%2d:%s (%s)\t%s\t%s" % (cnt,user.FullName,user.Handle,user.OnlineStatus,user.MoodText)
      cnt += 1
    self.printPrompt()
    print readline.get_line_buffer(),
    sys.stdout.flush()


  def OnFileTransferStatusChanged(self,file, status):
    self.printPrompt()
    if status==Skype4Py.fileTransferStatusPaused:
      if file.Type == Skype4Py.fileTransferTypeIncoming:
        print colored("Receiving file %s from %s paused.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))
      if file.Type == Skype4Py.fileTransferTypeOutgoing:
        print colored("Sending file %s to %s paused.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))

    if status==Skype4Py.fileTransferStatusRemotelyPaused:
      if file.Type == Skype4Py.fileTransferTypeIncoming:
        print colored("Receiving file %s from %s paused remotely.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))
      if file.Type == Skype4Py.fileTransferTypeOutgoing:
        print colored("Sending file %s to %s paused remotely.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))

    if status==Skype4Py.fileTransferStatusCancelled:
      if file.Type == Skype4Py.fileTransferTypeIncoming:
        print colored("Receiving file %s from %s canceled.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))
      if file.Type == Skype4Py.fileTransferTypeOutgoing:
        print colored("Sending file %s to %s canceled.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))

    if status==Skype4Py.fileTransferStatusCompleted:
      if file.Type == Skype4Py.fileTransferTypeIncoming:
        print colored("Receiving file %s from %s completed." ,'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))
      if file.Type == Skype4Py.fileTransferTypeOutgoing:
        print colored("Sending file %s to %s completed.",'green') %(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf_8'))

    if status==Skype4Py.fileTransferStatusFailed:
      if file.Type == Skype4Py.fileTransferTypeIncoming:
        error("Receiving file %s from %s failed."%(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf-8')))
      if file.Type == Skype4Py.fileTransferTypeOutgoing:
        error("Sending file %s to %s failed."%(file.FileName.decode('utf_8'), file.PartnerDisplayName.decode('utf-8')))
      if file.FailureReason == Skype4Py.fileTransferFailureReasonSenderNotAuthorized:
        error("Sender not authorized.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonRemotelyCancelled:
        error("Remotely cancelled.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonFailedRead:
        error("Failed read.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonFailedRemoteRead:
        error("Failed remote read.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonFailedWrite:
        error("Failed write.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonFailedRemoteWrite:
        error("Failed remote write.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonRemoteDoesNotSupportFT:
        error("Remote does not support file transfer.")
      if file.FailureReason == Skype4Py.fileTransferFailureReasonRemoteOfflineTooLong:
        error("Remote offline too long.")
    self.printPrompt()
    print readline.get_line_buffer(),
    sys.stdout.flush()

  def printPrompt(self):
    print "\x1b[2K",  #Clear current line
    if self.chat:
      print '\r(' + skype.chat.FriendlyName.encode('utf-8') + ')',
    print self.skype.CurrentUserProfile.FullName +  ':',
    sys.stdout.flush()



  #Print Error Message
  def error(self,mes):
    print colored("\x1b[2K[ERROR] " + mes,'red')


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
    print "\x1b[2K",  #Clear current line
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

    urlpattern = re.compile(r"""
    (
      (
        (
          [0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}|   #IPAdress
          (
            (
              ((news|telnet|nttp|file|http|ftp|https)://)|  #Scheme
              (www|ftp)[-A-Za-z0-9]*\.                      #HostName 
            )
            ([-A-Za-z0-9\.]+)                               #HostName 
          )
        )
        (:[0-9]*)?                                          #Port
      )
        (
          /[-A-Za-z0-9_\$\.\+\!\*\(\),;:@&=\?/~\#\%]*|      #Path
        )   
    )
    """,re.VERBOSE)

    urls = urlpattern.findall(mes)
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
