Hi!

First, thanks a heap for downloading my mailserver.

Please read this all before begining, it's sortof a 'quickstart' tutorial, and explains quite well what the server does, and how to do it.

Attached is a fully working opensource vb5/vb6 mailserver. It contains an smtp server, a pop3 server, and a very basic webmail site, all intergrated into one gui-less exe.

Before you even begin, you need to set one constant, otherwise it wont work. You need to tell the mailserver your dommain name. If you don't, it will not work, and will exhibit some REALLY strange behaviour, ie webmail wont work, cookies wont be set, mails will get into never ending loops, etc. It's the hostname constant in the functions module. this needs to be set to what follows the @ in your email messages. ie, if you ran hotmail, this'd be 'hotmail.com', if your on a little more of a budget then this, your machines ip address would do fine. Also, don't just think you can hijack hotmail by putting 'hotmail.com' in there, it needs to be the dommain name, or Ip address, registered to the machine on which it is running.

The next thing to do, restart your computer. This is optional, but it makes a HUGE difference in performance, ie being able to accept 30 silmultanius connections as opposed to 2. Also, try to shut down PWS, Real player, MSN, Trialian, Y!, ICQ, IE, OE, SETI@HOME, My webserver, UD, any adkillers, etc. that may be taking up memory, sockets, or performance on your machine.

After you do that, run the mailserver, hopefully it all should run fine (If your running another POP3, SMTP, or HTTP server it'll panic and shut down). You should do the following, in this order, to verify that it all works:

- Compile and Run the program. When you get it, in the queue will be a message, addressesed to me, saying 'It works'. I guess it isn't spyware if I tell you it's there. It tests the que mechanisms, and gives me a rough idea of how many people are using my mailserver. I get absolutely no personal information, I don't even get your email address. I just get 30 emails at the end of the day saying 'your mailserver worked'. If you don't want this to happen (and some people dont), clear the queue (\email\out) before you run it.

- Go into a web browser, and go to http://hostname (where hostname is what you entered as the constant before you even opened the peogram). You can do this and all steps from any computer on the same network as the mailserver. Signup for a new account. Enter an alternate email address, and an SMS address if you have a mobile phone (ie 0402145765@mobile.att.net), and a username/password. the sms & alt are optional, but useful. (new email alerts, etc.). The webmail is a redeye job (2AM to 5:30AM), so, it's very crap, and a little lagy (especially with the icons on the side). But, it's the only place you can signup on.

- go to your inbox (you will need to login), and read the message that was sent to you on signup. It contains your POP3 username, POP3 password, SMTP server, and POP3 server.

- Go into outlook, or whatever you use, and create a new account using those settings.

- Go 'check pop3 account' or 'recieve email' or whatever the menu item is, and outlook should log in, and download that first message in your inbox.

- Create a new email in outlook, and, making sure the right account is selected, email it to another account you have on some other server. (Be sure to check for any status emails from the server, ie, if you're not online, it will que it instead of sending it). ie, so it goes through MY smtp server, as oposed to your ISP's or your company's. And check to see whether it comes, it may take up to 30 seconds, depending on their host, your internet connection, and the size of the msg.

- Now, in outlook, create a new email to POP3USERNAME@HOSTNAME, (with POP3USERNAME and HOSTNAME being their values), see if that works.

- Check to see if it arived in your POP3 account, if you entered an SMS address, you should get a notification about now.

- Go to a site like hotmail or whatever, and send an email to POP3USERNAME@HOSTNAME. This step sometimes fails due to the ISP blocking SMTP connections (most ISP's do). The computer running the mailserver needs to accept incomming connections from the internet, ie, no proxys or anything, to do this, you need to be at the forefront of your internet connection, I'm sure you can set up your proxy server to route the connections through, just, don't ask me how.... :-)

- go back to outlook, and now try 'secure login' or whatever they call it. I can't get this to work from outlook 5, my version sends the wrong command (UIDL instead of APOP), but my server follows the protocol, and this works from telnet.

- Go to Planetsourcecode, and vote for this code.

And that's it! A complete mail server!

I wrote a very powerful webserver for the PSC code of the month competition, It was comming second by only one vote, with 4 days to go. Then, some idiot hacked the site and deleted some code. Inlcluding mine, the one above me, and the 4 or so below me. Ian declared the competition invalid, even thou I was able to give him the voting log for all of those projects that were lost within 3 hours of the deletion, he canceled the competition. (It's called 'A complete perl, + PHP and ASP webserver (be your own geocities)', check it out if you like.)

This mail server is basically me trying again, I really need some super qualification like "coder of the month" on my resume (I'm only 16), both for RAC and the rest of my employment life.

You are free to use this mailserver for whatever you want to, do anything you like to it, You can use it as the host for your million dollar business without any royalties or anything. I mean, the POP3 and SMTP server are reliable (webmail was thrown together between 2AM and 5:30AM Saturday morning before release, hence, a little buggy in spots.). You may wish to redo the 'accountsize' calculation, or get rid of it, as it slows down with large messages.

Also, I offer no warrenty whatsoever. I will take no blame for this mailserver being responsible for any loss of income, damage to property, random sensless killings, destruction of planets, death of patients on life support, SMSing one while in a movie and they in turn get killed by angry movie goers, space shuttles burning up on reentry, rising your IT staff's dental bills, or any other inconvienence to you. Basically, you use this, you must put up with the consequences yourself. I really can't afford the legal bills at the moment, so dont use it to maintain anything of value. (You get what you pay for, and you got this for free)

Happy Mailservering! If it works out for you, let me know! please! I dont have 4 IM apps and check my email 5 times a day for nothing!

Ashley - Sun 30 June 2002
--
Ashley___harris@hotmail.com
MSN: ashley___harris@hotmail.com
ICQ: 153577070
AIM: Ashley000Harris
Y!M: a_s_h_l_e_y_h_a_r_r_i_s