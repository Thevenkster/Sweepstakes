'''
This python script can be used to randomly assign teams to players using an
excel spreadsheet as input. Please ensure that the input files/sheets/columns
follow the naming conventions of the code to avoid errors/exceptions.

What this script does:
1. Opens an xlsx spreadsheet X
2. Reads the column with participants' & teams 'names & writes them to sets
3. Randomly picks a participant and assign both teams
4. Pops them from their sets
5. Saves the popped values into lists
6. Repeats until all sets are empty
7. Writes results to a new xlsx spreadsheet
'''
#importing libraries
import random
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#main function of the program
def main():

    #Read first sheet of the input file
    xl = pd.read_excel('inputdoc.xlsx', sheetname='Sheet1')

    #initialize sets for the players, pot1 teams & pot2 teams
    playerset = set()
    pot1set = set()
    pot2set = set()

    #Add all players, pot1 teams & pot2 teams to their respective set structures
    #We use sets instead of lists as they are unordered; hence more random
    for i in xl.index:
        playerset.add(xl['Player List'][i])
        pot1set.add(xl['Pot 1 List'][i])
        pot2set.add(xl['Pot 2 List'][i])

    #Create lists for assigning the teams to the players
    #We use lists here since the indicies will indicate who was assigned which teams
    players = []
    pot1 = []
    pot2 = []

    #Dataframe for output spreadsheet
    df = pd.DataFrame()

    #Writing into a new document instead of the old one for transparency
    writer = pd.ExcelWriter('AssignedTeams.xlsx', engine = 'xlsxwriter')

    #Loop through the index count
    for i in xl.index:

        #Choose a random element from each set
        chosenpl = random.choice(tuple(playerset))
        chosenpot1 = random.choice(tuple(pot1set))
        chosenpot2 = random.choice(tuple(pot2set))

        #Remove those elements from the set to avoid repetitions
        playerset.remove(chosenpl)
        pot1set.remove(chosenpot1)
        pot2set.remove(chosenpot2)

        #Write the selected values to their respective lists
        players.append(chosenpl)
        pot1.append(chosenpot1)
        pot2.append(chosenpot2)

    #Assign the list of values to their respective columns in the dataframe
    df['Player Name'] = players
    df['Team 1'] = pot1
    df['Team 2'] = pot2

    #Convert the dataframe to the excel format
    df.to_excel(writer, sheet_name = 'Sheet1')

    #Save the spreadsheet
    writer.save()

#Without this the code would execute even if the script was imported as a module
if __name__ == "__main__":
    main()
