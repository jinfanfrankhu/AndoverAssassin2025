import woof

'''
sendall(message, pth)

shuffle(pth)

emailAssignment(first, last, pth)

emailAllAssignments(pth)

kill(first, last, sendemail=Boolean, pth)
'''
# message = '''
# It goes without saying that Assassin is cancelled for the forseeable future.
# I'll be sending out more details when things settle down a bit. The loss was hard for me and I'm sure it was for you too. We can work through this together.
# '''
# woof.sendall(message, "killers.csv")

woof.shuffle("26.csv")
woof.shuffle("28.csv")

# woof.emailAllAssignments("killers.csv")

#woof.kill("Aeva", "Cleare", sendemail=True, pth="killers.csv")