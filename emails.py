file = open("email_address.txt", "w+")
massage = "sender's Email address: "

while True:
    senders_email = input(massage)
    if '@' in senders_email and " " not in senders_email:
        file.write(senders_email)
        break
    else:
        print('\nInvalid EMail address\n\n')

file.close()