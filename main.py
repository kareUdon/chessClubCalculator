import xlsxwriter # For writing to Excel
import os.path # To save the sheet to a path so it doesn't get lost

scores = {} # Scores will be in a dictionary 
howMany = int(input("How many players in ranking?")) # Gets number of players

def dictionaryOfScores():
  
  # Will require a bit of work in typing in the players' stats
  # Scores are calculated by adding wins and losses
  # Adding wins to the total wins and losses
  # Then subtracting the losses

  for i in range(0, howMany):

	  name = input("What's your name?")
	  wins = int(input("How many wins?"))
	  losses = int(input("How many losses?"))
	  score = str((wins + losses) + wins - losses) 
	
	  scores.update({name: score})

def writeToSheet():

	workbook = xlsxwriter.Workbook('data.xlsx')
	worksheet = workbook.add_worksheet()

	# Initial values of rows and columns
	row = 0
	col = 0

	# Getting the save path first as according to user preferences
	save_path = input("Please type in the filepath of where you want to store this.")

	for key in scores.keys(): #Iterates through keys and adds a row for each of them

		row += 1
		worksheet.write(row, col,key)

		for item in scores[key]: # This is for iterating through the values in the keys
			worksheet.write(row, col + 1, item)
			row += 1 

	workbook.close()

	location = os.path.join(save_path, 'data.xlsx')

dictionaryOfScores()
writeToSheet()

# Printed it just for testing purposes, xlsx file should be in specified save path
print(scores)
