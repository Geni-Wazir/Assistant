import pyttsx3
import time
import pygame
import subprocess
import webbrowser
import smtplib
import random
import speech_recognition as sr
import wikipedia
import datetime
import os
import sys
import re
from PIL import Image
import numpy as np
import imaplib
import email
import email.header
import cv2
import pickle



location = os.getcwd()

engine = pyttsx3.init('sapi5')


def speak(audio):
        engine.say(audio)
        engine.runAndWait()


def background_sound():
    os.chdir('background')
    pygame.mixer.init()
    pygame.mixer.music.set_volume(0.1)
    pygame.mixer.music.load('background.mp3')
    pygame.mixer.music.play(1000000000)


#=-----------------------------------------------------

internet_connection = subprocess.getoutput("ping google.com")
background_sound()
speak('Hello Sir, I am your digital assistant jiva!')
speak('checking for internet connection')
if internet_connection != "Ping request could not find host google.com. Please check the name and try again.":
    time.sleep(2)
    speak('you are connected to internet')
    speak('conneting to local servers')
    speak('conneting to data bases')
    speak('configuring system files')
    time.sleep(2)
    speak('connected to local servers')
    speak('connected to data bases')


    speak('system files configured properlly')

    speak('starting all services')

    for i in range(5):
        change_wallpaper = 'reg add "HKEY_CURRENT_USER\\Control Panel\\Desktop" /v Wallpaper /t REG_SZ /d {} /f'.format(os.getcwd()+'\\Wallpaper.png')
        subprocess.Popen(change_wallpaper, shell=True)
        subprocess.Popen('RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters', shell=True)
    os.chdir(location)
    subprocess.Popen('C:\\Program Files\\Rainmeter\\Rainmeter.exe', shell=True)
    time.sleep(2)
    speak('all services started properlly')
    speak("today is")
    speak(time.strftime("%x"))
    speak('current time ')
    speak(time.strftime("%X"))

    l = []
    count = 0

    def computer_specification():
        a = str(subprocess.getoutput('systeminfo'))
        b = a.find('System Model')
        c = a.find('OS Version')
        d = a.find('OS Manufacturer')
        e = a.find('System Type:               ')
        f = a.find('Processor')
        g = a.find("Registered Owner")
        h = a.find('Registered Organization')
        i = a.find('OS Name')
        j = a.find('System Manufacturer')
        k = a.find('Total Physical Memory')
        l = a.find('Available Physical Memory')

        speak(a[g:h-1]) # user [windows]
        speak(a[i:c-1]) # os name
        speak(a[c:d-1].replace('N/A', '')) # os version
        speak(a[e:f-1].replace('x', '')) # base [32 bits, 64 bits]
        speak(a[j:b-1]) # company
        speak(a[k:l-1].replace('Total Physical Memory', 'ram installed')) # ram



    def face_trainner():
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        image_dir = os.path.join(BASE_DIR, 'images')

        face_cascade = cv2.CascadeClassifier('cascades/haarcascade_frontalface_alt2.xml')
        recognizer = cv2.face.LBPHFaceRecognizer_create()


        current_id = 0
        label_ids = {}
        y_labels = []
        x_train = []

        for root, dirs, files in os.walk(image_dir):
            for file in files:
                if (file.endswith('.png') or file.endswith('.jpg')):
                    path = os.path.join(root, file)
                    label = os.path.basename(os.path.dirname(path)).replace(" ", "-").lower()
                    if not label in label_ids:
                        label_ids[label] = current_id
                        current_id += 1
                    id_ = label_ids[label]
                    pil_img =Image.open(path).convert('L')
                    size = (550,550)
                    final_image = pil_img.resize(size, Image.ANTIALIAS)
                    image_array = np.array(pil_img, "uint8")
                    faces = face_cascade.detectMultiScale(image_array, scaleFactor=1.1, minNeighbors=5)

                    for (x,y,w,h) in faces:
                        roi = image_array[y:y+h, x:x+w]
                        x_train.append(roi)
                        y_labels.append(id_)

        with open('label.pickle', 'wb') as f:
            pickle.dump(label_ids, f)

        recognizer.train(x_train, np.array(y_labels))
        recognizer.save('trainner.yml')



    def face_recognizer():
        face_cascade = cv2.CascadeClassifier('cascades/haarcascade_frontalface_alt2.xml')
        recognizer = cv2.face.LBPHFaceRecognizer_create()

        recognizer.read('trainner.yml')

        label ={"person name": 1}
        with open('label.pickle', 'rb') as f:
            label = pickle.load(f)
            label = {v:k for k,v in label.items()}

        cap = cv2.VideoCapture(0)

        speak('Sir please press Q to stop the Recognizer')
        while  True:
            rect, frame = cap.read()
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5)
            for (x,y,w,h) in faces:
        #print(x, y, w, h)
                roi_gray = gray[y:y+h, x:x+w]
                roi_color = frame[y:y+h, x:x+w]

            id_, conf = recognizer.predict(roi_gray)
            if conf >= 45 and conf <= 85:
                print(id_)
                print(label[id_])
                font = cv2.FONT_HERSHEY_SIMPLEX
                name = label[id_]
                color = (0,0,255)
                stroke = 1
                cv2.putText(frame, name, (x,y), font, 1, color, stroke, cv2.LINE_AA)

            color = (0, 255, 0)
            stroke = 2
            end_chord_x = x + w
            end_chord_y = x + h
            cv2.rectangle(frame, (x,y),(end_chord_x,end_chord_y), color, stroke)
            cv2.imshow('frame', frame)
            if cv2.waitKey(20) & 0xFF ==ord('q'):
                break


        cap.release()
        cv2.destroyAllWindows()




    def audio():
        directory = []
        files = []
        drives = []
        a = subprocess.getoutput("wmic logicaldisk get name").split("\n")[2:-3]
        for letter in range(0, len(a), 2):
            drives.append(a[letter].strip())
        drives = drives[1:-1]

        for drive in range(len(drives)):
            drives[drive] = drives[drive] + "\\"

        def to_get_all_directories():
            global i, a, k
            for i in drives:
                os.chdir(i)
                a = subprocess.getoutput("dir *.* /S /p")  # to get the output of the command that search the
            # files in a particular directory
                b = a.split("\n\n")  # to split the result according to the given instruction
                for k in b:
                    if " Directory of " in k:
                        directory.append(k[14:])   # take the path of the directories from the result

        to_get_all_directories()
        for data in directory:
            files.append(os.listdir(data))

        def execution_of_program():
            global audio_file, a, k
            audio_file = []
            a = 0
            k = 1
            for p in directory:
                n = 0
                for s in range(len(files[a])):
                    full_location = p + "\\" + files[a][n]
                    if full_location.endswith('mp3') or full_location.endswith('.m4a'):
                        audio_file.append(full_location)
                        k += 1
                    n += 1
                a = a+1
        execution_of_program()

        return random.choice(audio_file), k


    def video():
        directory = []
        files = []
        drives = []
        a = subprocess.getoutput("wmic logicaldisk get name").split("\n")[2:-3]
        for letter in range(0, len(a), 2):
            drives.append(a[letter].strip())
        drives = drives[1:-1]

        for drive in range(len(drives)):
            drives[drive] = drives[drive] + "\\"

        def to_get_all_directories():
            global i, a, k
            for i in drives:
                os.chdir(i)
                a = subprocess.getoutput("dir *.* /S /p")  # to get the output of the command that search the
            # files in a particular directory
                b = a.split("\n\n")  # to split the result according to the given instruction
                for k in b:
                    if " Directory of " in k:
                        directory.append(k[14:])   # take the path of the directories from the result

        to_get_all_directories()
        for data in directory:
            files.append(os.listdir(data))

        def execution_of_program():
            global video_file, a, k
            video_file = []
            a = 0
            k = 1
            for p in directory:
                n = 0
                for s in range(len(files[a])):
                    full_location = p + "\\" + files[a][n]
                    if full_location.endswith('.mp4') or full_location.endswith('.avi') or full_location.endswith('.wmv') or full_location.endswith('.mov'):
                        video_file.append(full_location)
                        k += 1
                    n += 1
                a = a+1
        execution_of_program()

        return random.choice(video_file), k



    def photo():
        directory = []
        files = []
        drives = []
        a = subprocess.getoutput("wmic logicaldisk get name").split("\n")[2:-3]
        for letter in range(0, len(a), 2):
            drives.append(a[letter].strip())
        drives = drives[1:-1]

        for drive in range(len(drives)):
            drives[drive] = drives[drive] + "\\"

        def to_get_all_directories():
            global i, a, k
            for i in drives:
                os.chdir(i)
                a = subprocess.getoutput("dir *.* /S /p")  # to get the output of the command that search the
            # files in a particular directory
                b = a.split("\n\n")  # to split the result according to the given instruction
                for k in b:
                    if " Directory of " in k:
                        directory.append(k[14:])   # take the path of the directories from the result

        to_get_all_directories()
        for data in directory:
            files.append(os.listdir(data))

        def execution_of_program():
            global photo_file, a, k
            photo_file = []
            a = 0
            k = 1
            for p in directory:
                n = 0
                for s in range(len(files[a])):
                    full_location = p + "\\" + files[a][n]
                    if full_location.endswith('.png') or full_location.endswith('.jpg') or full_location.endswith('.jpeg'):
                        photo_file.append(full_location)
                        k += 1
                    n += 1
                a = a+1
        execution_of_program()
        return photo_file, k

    def greetMe():
        currentH = int(datetime.datetime.now().hour)
        if currentH >= 0 and currentH < 12:
            speak('Good Morning Sir !')

        if currentH >= 12 and currentH < 17:
            speak('Good Afternoon Sir!')

        if currentH >= 17 and currentH !=0:
            speak('Good Evening Sir!')

    greetMe()


    speak('how may I help you sir.........')

    def myCommand():

        r = sr.Recognizer()
        global count
        with sr.Microphone() as source:
            if count ==0:
                count=1
                speak('Ready for your command sir')
            audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')

        except sr.UnknownValueError:
            speak('Sorry sir! I didn\'t get that!')
            speak('Please try once again littel louder and clear')
            query = myCommand()
        background_sound()
        return query



    def video_recorder():

        subprocess.call('mkdir Recorded_videos', shell=True)
        os.chdir('Recorded_videos')
        file = str(time.strftime("%x")) + str(time.strftime("%X")) + '.mp4'
        file = file.replace('/', '-')
        file = file.replace(':', '-')
        fps = 24.0
        myres = '720p'

        def change_res(cap, w ,h):
            cap.set(3, w)
            cap.set(4, h)

        std_demission = {
            '420p': (640, 480),
            '720p': (1280, 720),
            '1080p': (1920 ,1080),
            '4k': (3840,2160)
        }


        def get_dem(cap, res='720p'):

            w, h = std_demission['720p']
            if res in std_demission:
                w, h = std_demission[res]
            change_res(cap, w, h)
            return w, h

        video_type = {

            'avi': cv2.VideoWriter_fourcc(*'XVID'),
            'mp4': cv2.VideoWriter_fourcc(*'XVID')
        }

        def get_video_type(file):
            file , ext = os.path.splitext(file)
            if ext in video_type:
                return video_type[ext]
            return video_type['mp4']

        cap = cv2.VideoCapture(0)
        dim = get_dem(cap, res=myres)

        video_type_cv2 = get_video_type(file)

        out = cv2.VideoWriter(file, video_type_cv2, fps, dim)

        while  True:
            rect, frame = cap.read()
            out.write(frame)
            cv2.imshow('RECORDER', frame)
            if cv2.waitKey(5) & 0xFF ==ord('q') or cv2.waitKey(20) & 0xFF ==ord('Q'):
                break

        cap.release()
        out.release()
        cv2.destroyAllWindows()

        os.chdir(location)



    if __name__ == '__main__':

        while True:

            #query='read mail'
            query = myCommand()
            query = query.lower()

            if 'open youtube' in query:
                speak('okay Sir opening youtube for you')
                webbrowser.open('www.youtube.com')
                time.sleep(3)

            elif 'open google' in query:
                speak('okay Sir opening google for you')
                webbrowser.open('www.google.co.in')
                time.sleep(3)

            elif 'open facebook' in query:
                speak('okay Sir opening facebook for you')
                webbrowser.open('www.facebook.com')
                time.sleep(3)

            elif 'open instagram' in query:
                speak('okay Sir opening google for you')
                webbrowser.open('www.instagram.com')
                time.sleep(3)

            elif 'open gmail' in query:
                speak('okay Sir opening gmail for you')
                webbrowser.open('www.gmail.com')
                time.sleep(3)

            elif "open what'sapp" in query or "what's app" in query or 'whatsapp' in query:
                speak('okay Sir opening whatsapp for you')
                webbrowser.open('web.whatsapp.com')
                time.sleep(3)

            elif 'search' in query or 'find' in query:
                query = query.replace("search", "")
                query = query.replace('find', '')
                speak('Searching')
                speak(query)
                webbrowser.open(query)
                time.sleep(3)

            elif 'what' in query and 'time' in query:
                speak('time is')
                speak(time.strftime("%X"))

            elif 'what' in query and 'date' in query:
                speak("today's date")
                speak(time.strftime("%x"))

            elif 'joke' in query:
                file = open("jokes.txt", 'r')
                jokes = file.read()
                k = jokes.split('\n\n')
                speak(random.choice(k))

            elif 'play music' in query or 'play audio' in query  or 'play audios' in query:

                file, no = audio()
                command = '"' + file + '"'
                file = file.replace('_', "")
                statement = 'sir you have ' + str(no) + 'audio files \n and playing ' + file
                speak(statement)
                subprocess.call(command, shell=True)


            elif 'play movie' in query or 'play video' in query or 'play videos' in query:
                file, no = video()
                command = '"' + file + '"'
                file = file.replace('_', "")
                statement = 'sir you have ' + str(no) + 'video files \n and playing ' + file
                speak(statement)
                subprocess.call(command, shell=True)

            elif 'open photo' in query or 'open photos' in query or 'play photos' in query or 'play photo' in query:
                file, no = photo()
                random_photo = []
                if no > 20:
                    statement = 'sir you have '+ str(no) + 'photo \n and playing any 20 photo from them'
                    speak(statement)
                    for i in range(20):
                        photo = random.choice(file)
                        random_photo.append(photo)
                    for photos in random_photo:
                        img = cv2.imread(photos, 1)
                        resize = cv2.resize(img, (500,600))
                        cv2.imshow(photos, resize)
                        cv2.waitKey(3000)
                        cv2.destroyAllWindows()
                    time.sleep(3)

                else:
                    statement = 'sir you have ' + str(no) + 'photo \n and playing all of them'
                    speak(statement)
                    for photos in file:
                            img = cv2.imread(photos, 1)
                            resize = cv2.resize(img, (500,600))
                            cv2.imshow(photos, resize)
                            cv2.waitKey(3000)
                            cv2.destroyAllWindows()
                    time.sleep(3)

            elif 'save' in query:
                speak('what will be the name of the txt file')
                file_name = myCommand()
                speak('creating new file')
                speak('Sir the file has been created what to write in the file')
                content = myCommand()
                with open(file_name, 'w') as f:
                    f.write(content)

            elif ('click' in query or 'take' in query) and ('photo' in query or 'photos' in query):
                speak('Sir please take your position')
                location = os.getcwd()
                face_cascade = cv2.CascadeClassifier('cascades/haarcascade_frontalface_alt2.xml')
                subprocess.call('mkdir photo', shell=True)
                os.chdir('photo')
                cap = cv2.VideoCapture(0)
                rect, frame = cap.read()
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5)
                file_name = str(time.strftime("%x")) + str(time.strftime("%X")) + '.png'
                file_name = file_name.replace('/', '-')
                file_name = file_name.replace(':', '-')
                cv2.imwrite(file_name, frame)
                cv2.imshow('CAMERA', frame)
                cv2.waitKey(1500)
                cap.release()
                cv2.destroyAllWindows()

                os.chdir(location)


            elif 'face' in query and 'trainner' in query:

                speak('Starting face trainner sir')
                speak('trainning is in process please wait sir')
                face_trainner()
                speak('Sir trainning has been completed successfully')


            elif 'start' in query and 'recognizer' in query:

                speak('starting Recognizer')
                face_recognizer()

            elif 'start' in query and 'video' in query and 'recording' in query or 'recorder' in query:
                video_recorder()

            elif 'open my computer' in query or 'open this pc' in query:
                subprocess.call('start D:', shell=True)

            elif 'about you' in query or 'specification' in query:
                computer_specification()

            elif "what\'s up" in query or 'how are you' in query:
                stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
                speak(random.choice(stMsgs))

            elif 'read' in query and 'mail' in query:
                EMAIL_ACCOUNT = "example@gmail.com"       #Email Address
                PASSWD = "Password"              #Email Password
                EMAIL_FOLDER = "INBOX"  #Folder to be read

                def process_mailbox(M):
                    countIT = 0
                    rv, data = M.search(None, "ALL")    #UNSEEN if you want to see unread messages
                    if rv != 'OK':
                        speak("No messages found!")
                        return

                    for num in reversed(data[0].split()):
                        countIT = countIT+1
                        if (countIT == 6):
                            return
                        rv, data = M.fetch(num, '(RFC822)')
                        if rv != 'OK':
                            speak("ERROR getting message")
                            return

                        speak('sir  ')
                        msg = email.message_from_bytes(data[0][1])
                        temp = 'Message by %s saying %s' % (msg['From'], msg['Subject'])
                        line = re.sub('<[^>]+>', '', temp)      #Remove if you want it to say the email address too
                        speak(line)

                M = imaplib.IMAP4_SSL('imap.gmail.com') #Mail Server

                try:
                    rv, data = M.login(EMAIL_ACCOUNT, PASSWD)
                except imaplib.IMAP4.error:
                    speak("LOGIN FAILED!!! ")
                    sys.exit(1)

                rv, data = M.select(EMAIL_FOLDER)
                if rv == 'OK':
                    process_mailbox(M)
                    M.close()
                else:
                    speak("ERROR: Unable to open mailbox ")

                M.logout()

            elif 'have' in query and 'mail' in query:
                speak('checking for new mails')
                EMAIL_ACCOUNT = "example@gmail.com"       #Email Address
                PASSWD = "Password"              #Email Password
                EMAIL_FOLDER = "INBOX"  #Folder to be read

                def process_mailbox(M):
                    countIT = 0
                    rv, data = M.search(None, "ALL")    #UNSEEN if you want to see unread messages
                    if rv != 'OK':
                            speak("No messages found!")
                            return

                    for num in reversed(data[0].split()):
                            countIT = countIT+1
                            if (countIT == 6):
                                    return
                            rv, data = M.fetch(num, '(RFC822)')
                            if rv != 'OK':
                                    speak("ERROR getting message")
                                    return
                            msg = email.message_from_bytes(data[0][1])
                            temp = msg['From']
                            l.append(re.sub('<[^>]+>', '', temp))


                M = imaplib.IMAP4_SSL('imap.gmail.com') #Mail Server

                try:
                    rv, data = M.login(EMAIL_ACCOUNT, PASSWD)
                except imaplib.IMAP4.error:
                    speak("LOGIN FAILED!!! ")
                    sys.exit(1)

                rv, data = M.select(EMAIL_FOLDER)
                if rv == 'OK':
                    process_mailbox(M)
                    M.close()
                else:
                    speak("ERROR: Unable to open mailbox ")

                M.logout()
                if len(l) == 0:
                    speak("Sorry sir you don't have any new mail")
                else:
                    speak('sir you have mails from')
                    if len(set(l)) == 1:
                        speak(m)
                    else:
                        for m in set(l):
                            if m ==l[-1]:
                                speak('and')
                            speak(m)


            elif 'send email' in query:
                speak('Who is the recipient? ')
                recipient = myCommand()

                if 'me' in recipient or 'myself' in recipient:
                    try:
                        speak('What should I say? ')
                        content = myCommand()

                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.ehlo()
                        server.starttls()
                        server.login("example@gmail.com", 'Password')
                        server.sendmail('example@gmail.com', "example2@gmail.com", content)
                        server.close()
                        speak('Email sent!')

                    except:
                        speak('Sorry Sir! I am unable to send your message at this moment!')
                else:
                    speak('please enter the email address of the recipiant')
                    subprocess.call('python emails.py', shell=True)
                    file = open("email_address.txt", "r")
                    reciver_email = file.read()

                    try:
                        speak('What should I say? ')
                        content = myCommand()

                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.ehlo()
                        server.starttls()
                        server.login("example@gmail.com", 'Password')
                        server.sendmail('example@gmail.com', reciver_email, content)
                        server.close()
                        speak('Email sent!')

                        sent_email = open('Sent_Emails.txt', "a+")
                        from_email = 'From: abcd67649@gmail.com'
                        to_email = 'To: ' + reciver_email
                        content_of_the_email = 'content: ' + content
                        sent_email.write(from_email)
                        sent_email.write('\n')
                        sent_email.write(to_email)
                        sent_email('\n')
                        sent_email.write(content_of_the_email)

                    except:
                        speak('Sorry Sir! I am unable to send your message at this moment!')


            elif 'abort' in query or 'shutdown ' in query:
                speak('shuting down the System')
                speak('disconneting from local servers')
                speak('disconneting to data bases')
                speak('stopping all services')
                output = subprocess.getoutput('taskkill /im Rainmeter.exe /f')
                if 'not found' not in output:
                    speak('could not stop all services properlly')
                sys.exit()

            elif 'stop' in query or 'terminate' in query:
                stopwords = ['stop', 'terminate']
                querywords = query.split()
                resultwords  = [word for word in querywords if word.lower() not in stopwords]
                result = ' '.join(resultwords)
                command = 'taskkill /im ' + result + '.exe /f'
                l = subprocess.getoutput(command)
                if 'not found' not in l:
                    statement = 'stopping ' + result
                    speak(statement)
                else:
                    statement = result + ' not running'
                    speak(statement)


            elif 'shutdown' in query and 'pc' in query:
                subprocess.call('shutdown /p', shell=True)

            elif 'restart' in query and 'pc' in query:
                subprocess.call('shutdown /r', shell=True)

            elif 'log off' in query or 'logoff' in query and 'pc' in query:
                subprocess.call('shutdown /l', shell=True)

            elif 'hello' in query:
                speak('Hello Sir')

            elif 'activate sleep mode' in query:
                speak('Activating sleep mode')
                time.sleep(2)
                speak('Sleep mode activated for 5 minutes')
                time.sleep(300)
                speak('Deactivating sleep mode')
                speak('back to the work')

            else:
                query = query
                speak('Searching...')
                try:
                    results = wikipedia.summary(query, sentences=2)
                    speak('According to wikipedia - ')
                    speak(results)

                except:
                    speak("SORRY! sir didn't found relevent answer in my data base")
                    webbrowser.open('www.google.com')
                    speak("searching it on  google for you")

else:
    speak('sir you are not connected to internet')
    speak('can not connect to the servers')
    speak('can not connect to the servers')
    speak('please check your internet connection')
