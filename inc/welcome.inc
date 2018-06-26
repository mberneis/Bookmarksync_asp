<%
function welcometxt (user,id)
 Dim body

 body = "Dear " & user &  vbCrLf
 body = body & " " & vbCrLf
'		012345678901234567890123456789012345678901234567890123456789012345678901
 body = body & "Welcome to SyncIT.com.  We thank you for your registration, and I"	& vbCrLf
 body = body & "would like to alert you to the services you are eligible for as a"	& vbCrLf
 body = body & "SyncIT customer. "							& vbCrLf
 body = body & "First, very importantly, to activate your registration, please click "	& vbCrLf
 body = body & "      http://syncit.com/common/verify.asp?id=" & id			& vbCrLf
 body = body                                                                            & vbCrLf
 body = body & "SyncIT services you will want to check out include:"		& vbCrLf
 body = body & "BookmarkSync."                                                          & vbCrLf
 body = body & "The #1 ranked storage and synchronization service for bookmarks,"	& vbCrLf
 body = body & "favorites and favorite places. No need for extra work importing and"	& vbCrLf
 body = body & "exporting bookmark files - BookmarkSync keeps track of changes you"	& vbCrLf
 body = body & "make on your browser.  For access from any computer you use on a"	& vbCrLf
 body = body & "regular basis, download our BookmarkSync utility, and click on"	& vbCrLf
 body = body & "the SyncIT icon in the lower right-hand corner of your screen.  For"	& vbCrLf
 body = body & "access from a new computer, log in with your user name and password"	& vbCrLf
 body = body & "on the www.syncit.com web site for easy access to your bookmarks and"	& vbCrLf
 body = body & "favorites."								& vbCrLf
 body = body                                                                            & vbCrLf
 body = body & "MySync Collections."							& vbCrLf
 body = body & "Publish your bookmark collections on any subject and invite your"	& vbCrLf
 body = body & "friends and colleagues to subscribe to your links, or mark them as"	& vbCrLf
 body = body & "'public' and invite the entire SyncIT community to subscribe.  Check"	& vbCrLf
 body = body & "out some of our members' cool public collections."			& vbCrLf
 body = body & "      http://syncit.com/collections/"					& vbCrLf
 body = body                                                                            & vbCrLf
 body = body & "QuickSync News."							& vbCrLf
 body = body & "Sign up for live updates from magazines and subjects of your choice,"	& vbCrLf
 body = body & "delivered direct to your SyncIT icon.  Just go to www.syncit.com and"	& vbCrLf
 body = body & "click QuickSync News."							& vbCrLf
 body = body                                                                            & vbCrLf
 body = body & "SyncIT Express."							& vbCrLf
 body = body & "Sending files too large for your current e-mail service?"	& vbCrLf
 body = body & "Do you need confirmation when your delivery is received?"	& vbCrLf
 body = body & "SyncIT Express is your universal courier!"							& vbCrLf
 body = body                                                                            & vbCrLf
 body = body & "Again, thank you for joining SyncIT.  We'll do our best to make you a"	& vbCrLf
 body = body & "happy customer - please write to me and tell me how we're doing.  If"	& vbCrLf
 body = body & "you're happy with our service, please tell your friends to check it"	& vbCrLf
 body = body & "out - we greatly appreciate your referrals."				& vbCrLf
 body = body & "You have a 30 day free trial to appreciate this outstanding utility." & vbCrLf
 body = body & "If by this time you are convinced of the quality and usefulness of our " & vbCrLf
 body = body & "product, please purchase a license for the OS you are using." & vbCrLf
 body = body                         & vbCrLf
 body = body & "Regards,"            & vbCrLf
 body = body                         & vbCrLf
 body = body & "Michael Berneis"     & vbCrLf
 body = body & "CEO"                 & vbCrLf
 body = body & "SyncIT.com"          & vbCrLf
 body = body                         & vbCrLf
 body = body & "michaelb@syncit.com" & vbCrLf
 body = body & "www.syncit.com"      & vbCrLf
welcometxt = body
end function
%>
