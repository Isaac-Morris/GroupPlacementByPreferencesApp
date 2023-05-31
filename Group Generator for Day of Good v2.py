# NOTE: For this to work, you should probably ensure that the signups data is consistent with the list of charities, and use Excel find and replace
# if the charity names aren't consistent. For example, if you change the Google Form after the signups have opened, it won't change every sign up who
# previously clicked that option, only the future ones, so you'll need to make it consistent

# IMPORTS ----------------------------------------------------------------------------------

# NOTE on pandas for new users: If assigning, you need to use df.at[i, j] or df.loc[i, j] instead of just df[i, j], due to https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy. .at is preferred for single variable assigning
import pandas as pd
from collections import defaultdict, Counter
from itertools import chain
from pathlib import Path

# CONSTANTS --------------------------------------------------------------------------------

# These could probably be read from an excel sheet or inputted, if you wanted to make a more easy-to-use interactive wrapper, but I cbs
PATH = Path(__file__).parent
PLACES_FILEPATH = PATH.joinpath("places_input.xlsx")
SIGNUPS_FILEPATH = PATH.joinpath("signups_input.xlsx")
MAX_NUM_EMAILS = 3
MAX_RECURSIONS = 50

# PREPROCESSING ---------------------------------------------------------------------------

# Checks for emails that appear multiple times, treats emails as unique IDs.
# These currently need to be dealt with manually by the user
def check_repeated_emails(df_signups):
    line_ids_for_emails = defaultdict(list)

    for i in range(len(df_signups)): # Record the line numbers where each email appears
        for col_name in get_email_col_names():
            email = str(df_signups.at[i, col_name]).strip().lower() # To avoid errors due to different formatting or whitespace
            line_ids_for_emails[email].append(i + 2) # +2 offset to account for header line

    if "nan" in line_ids_for_emails.keys():
        line_ids_for_emails.pop("nan") # For empty cells
    
    line_ids_for_emails = line_ids_for_emails.items()
    print("There are " + str(len(line_ids_for_emails)) + " unique emails who have signed up")

    # Get only the emails which have been repeated
    repeated_emails = [x for x in line_ids_for_emails if len(x[1]) > 1]

    if len(repeated_emails) > 0:
        print("There are " + str(len(repeated_emails)) + " emails which appear in the signup form multiple times. \nThese may be groups who signed up multiple times (once for each person), or a person who changed their response after submitting. Generally, you can simply delete the earliest instance where they appear, as the later one is probably a resubmission with more accurate details\nPlease see the output excel for a list of these emails and amend before rerunning the code")
        print("\n")
        repeated_emails = pd.DataFrame(repeated_emails) 
        repeated_emails.columns = ["email","appears on lines"]
    else:
        repeated_emails = pd.DataFrame()

    return repeated_emails

# A functions which checks for commas in the names of charities and removes them. It then updates all the signups in df_signups
# to remove commas in the charity name, so the string of charities they signed up for can be split into a list
def remove_commas(df_charities, df_signups):
    charities_with_commas = []

    for i in range(len(df_charities)):
        charity_name = df_charities.at[i, 'charity']
        if ',' in charity_name:
            charities_with_commas.append([charity_name, charity_name.replace(',', '')])
            df_charities.at[i, 'charity'] = charity_name.replace(',', '')

    for i in range(len(df_signups)):
        for charity_replacement in charities_with_commas:
            df_signups.at[i, 'chosen_charities'] = df_signups.at[i, 'chosen_charities'].replace(charity_replacement[0], charity_replacement[1])


# Counts the number of participants in each group (line of the dataframe)
def count_num_participants_in_group(df_signups):
    # Initialise columns
    df_signups['num_participants'] = 0
    df_signups['participant_emails'] = [[] for i in range(len(df_signups))]
    
    for i in range(len(df_signups)): # Create a list of participants in line and count length
        for col_name in get_email_col_names():
            if type(df_signups[col_name][i]) != float:
                df_signups.loc[i, 'participant_emails'].append(df_signups[col_name][i])
        df_signups.loc[i, 'num_participants'] = len(df_signups.loc[i, 'participant_emails'])

# Counts how many charities each group chose
def count_charities_picked_by_group(df_signups):
    # Initialise columns
    df_signups['num_charities_picked'] = 0
    df_signups['charity_list'] = [[] for i in range(len(df_signups))]

    for i in range(len(df_signups)):
        df_signups.at[i, 'charity_list'] = df_signups.at[i, 'chosen_charities'].split(', ')
        df_signups.at[i, 'num_charities_picked'] = len(df_signups.loc[i, 'charity_list'])


def count_num_time_charity_picked(df_charities, df_signups):
    # Get the counts of each time a charity appears
    all_charities_picked = df_signups['charity_list']
    charity_counts = Counter(chain(*list(all_charities_picked)))

    df_charities['total_signups'] = 0
    
    for i in range(len(df_charities)):
        charity_name = df_charities.at[i, 'charity']
        count = charity_counts[charity_name]
        del charity_counts[charity_name]
        df_charities.at[i, 'total_signups'] = count

    print("The following places were in the signups, but did not appear in the input file (" + str(PLACES_FILEPATH) + "):")
    for pair in charity_counts.items():
        print("    -" + pair[0] + "     appeared " + str(pair[1]) + " times")
    print("This may be because you changed the name of the option on the Google Form after people had signed up to it. If this is the case, please use CTRL-F on one of the Excel documents and replace all instances of the previous option name to the updated name so they're consistent")
    print("Otherwise, if you were expecting for an option to not be included (because that option is no longer available but students already signed up for it, or because you're assigning people there manually), you can ignore it in this list")
    print("\n")
    

# For each charity, the ratio of signups : max possible places is what we use for sorting, here labelled as "ratio" 
def get_charity_ratios(df_charities, df_signups):
    count_num_time_charity_picked(df_charities, df_signups)
    df_charities['ratio'] = df_charities.total_signups / df_charities.max_number
    
# Sorts charities so we deprioritise charities who had significantly more interest than possible places to fill
def sort_charities_by_ratio(df_charities, df_signups):
    df_charities.sort_values('ratio', inplace = True)
    df_charities.reset_index(inplace = True, drop = True)
    
# Sort participants so we prioritise placing large groups and groups who didn't select many charities. 
# While this does reward people who made our life difficult and only selected one/two places, it also means we're
# more likely to place everyone
def sort_participants_by_group_size_and_num_charities_selected(df_signups):
    df_signups.sort_values(['num_participants', 'num_charities_picked'], ascending = [False, True], inplace = True)
    df_signups.reset_index(inplace = True, drop = True)


# PROCESSING (PLACING PARTICIPANTS) -----------------------------------------------------------------------------

# This actually does the process of assigning our participants to charities
def assign_participants_to_charities(df_charities, df_signups):
    # Initialising lists of where people have been assigned
    df_charities['participant_count'] = 0
    df_charities['participant_ids'] = [[] for i in range(len(df_charities))]

    # Check that every person has preferenced at least on charity which actually exists, and create a list of all the people who couldn't be assigned
    ids_with_no_existing_options = []
    all_existing_charities = set(df_charities['charity'])
    too_late_signups_list = []

    for i in range(len(df_signups)):
        overlap = all_existing_charities & set(df_signups.loc[i, 'charity_list'])
        if len(overlap) == 0: # Either this person preferenced no charities or the charities they preferenced don't exist in the charity list
            ids_with_no_existing_options.append(i)
        else:# Otherwise, this person has valid options so assign them
            assign_to_charity(i, df_signups, df_charities, too_late_signups_list = too_late_signups_list)

    df_unassignable_people = generate_dataframe_of_unassignable_people(too_late_signups_list, ids_with_no_existing_options, df_signups)

    return df_unassignable_people


# We can assume here that everyone in this has at least one charity they can feasibly be added to
# Add each signup to a charity that's currently got space for them
def assign_to_charity(participant_id, df_signups, df_charities, index_to_skip = -1, num_recursions = 0, too_late_signups_list = []):
    # We first try attempting to find a charity with less than the minimum required numbers
    if assign_to_charities_below_min_numbers(participant_id, df_signups, df_charities):
        return
            
    # If every charity for this person has their minimum required numbers, we will then try to fill to maximum (basically same process as above)
    if assign_to_charities_between_min_and_max_numbers(participant_id, df_signups, df_charities):
        return

    # If that's not possible, we need to check if this recursion has just been going forever. If it has, we flag the participants as needing to be added manually
    if num_recursions > MAX_RECURSIONS:
        find_and_replace_latest_signups(participant_id, df_signups, df_charities, too_late_signups_list)
        return

    # There's no charity which currently has space, so we'll go the one which has the smallest signup:max ratio, and force the previously added person to move
    if assign_to_charity_and_move_most_recent_addition(participant_id, df_signups, df_charities, index_to_skip, num_recursions, too_late_signups_list):
        return

    # If we get to this point without being able to put the participant anywhere, either we have a problem or we need to put them in the charity we bumped them from
    if index_to_skip == -1:
        print("There's an error somewhere in the participant assigning code lol. This is probably because there are places which are in signups but not in the input file")
        # print(df_signups.loc[participant_id, :])
        # print("Recursion num:" + str(num_recursions))
    else:
        # Since we haven't been able to assign this participant to any other charity than the one they were originally bumped from, we'll assign them there and recurse
        assign_to_previously_bumped_from_charity(participant_id, df_signups, df_charities, index_to_skip, num_recursions, too_late_signups_list)
        return

            

# Returns true if there was a charity below the minimum numbers to add the participant to
def assign_to_charities_below_min_numbers(participant_id, df_signups, df_charities):
    participant_charity_list = df_signups.loc[participant_id, 'charity_list']
    num_participants_in_group = df_signups.loc[participant_id, 'num_participants']

    for charity_index in range(len(df_charities)):
        if (df_charities.loc[charity_index, 'charity'] in participant_charity_list):
            if (df_charities.loc[charity_index, 'participant_count'] + num_participants_in_group) <= df_charities.loc[charity_index, 'min_number']:
                # There's space in this preference for the group/participant. Add them
                add_participant_to_charity(participant_id, charity_index, df_signups, df_charities)
                # print("Success for ID " + str(i) + " assigned to " + df_charities.loc[charity_index, 'charity'])
                return True
    
    return False

# Returns true if there was a charity which is between the minimum and maximum numbers when a participant/group is added
def assign_to_charities_between_min_and_max_numbers(participant_id, df_signups, df_charities):
    participant_charity_list = df_signups.loc[participant_id, 'charity_list']
    num_participants_in_group = df_signups.loc[participant_id, 'num_participants']

    for charity_index in range(len(df_charities)):
        if (df_charities.loc[charity_index, 'charity'] in participant_charity_list):
            if (df_charities.loc[charity_index, 'participant_count'] + num_participants_in_group) <= df_charities.loc[charity_index, 'max_number']:
                add_participant_to_charity(participant_id, charity_index, df_signups, df_charities)
                # print("Success for ID " + str(i) + " assigned to " + df_charities.loc[charity_index, 'charity'])
                return True
            
    return False

# We've exceeded our maximum number of recursions, so we'll place this participant in the list, and then bump off the participants who signed up too late (by timestamp)
def find_and_replace_latest_signups(participant_id, df_signups, df_charities, too_late_signups_list):
    participant_charity_list = df_signups.loc[participant_id, 'charity_list']
    num_participants_in_group = df_signups.loc[participant_id, 'num_participants']

    # We'll add this participant to the first possible charity in the list, because chances are that's where the problem is lol
    for charity_index in range(len(df_charities)):
        if (df_charities.loc[charity_index, 'charity'] in participant_charity_list):
            
            add_participant_to_charity(participant_id, charity_index, df_signups, df_charities)

            df_charities.at[charity_index, 'participant_ids'] = sorted(df_charities.at[charity_index, 'participant_ids'], key = lambda id: df_signups.at[id, 'timestamp']) # I probs need to check this works lol
            
            # Note the participant/s who signed up late and need to be bumped off so this charity can be below the minimum numbers
            while df_charities.at[charity_index, 'participant_count'] > df_charities.at[charity_index, 'max_number']:
                last_signup_id = pop_most_recently_added_participant(charity_index, df_signups, df_charities)
                too_late_signups_list.append(last_signup_id)
            
            return

# Returns true if there is somewhere this participant can be added that is NOT where they were previously bumped from (taking the place of whoever was most recently added)
def assign_to_charity_and_move_most_recent_addition(participant_id, df_signups, df_charities, index_to_skip, num_recursions, too_late_signups_list):
    participant_charity_list = df_signups.loc[participant_id, 'charity_list']
    num_participants_in_group = df_signups.loc[participant_id, 'num_participants']

    for charity_index in range(len(df_charities)):
        if (df_charities.loc[charity_index, 'charity'] in participant_charity_list) and (charity_index != index_to_skip):
        
            # Bump off the most recent addition to this group and force them to go to a different charity lol
            most_recent_addition_id = pop_most_recently_added_participant(charity_index, df_signups, df_charities)

            prepend_participant_to_charity(participant_id, charity_index, df_signups, df_charities) # Prepend the ID (to make infinite loops only occur after checking the whole list)
            # print ("ID " + str(i) + " assigned to " + df_charities.loc[charity_index, 'charity'])

            # Assign the popped id to a charity (if possible, not this one which they got popped from)
            assign_to_charity(most_recent_addition_id, df_signups, df_charities, index_to_skip = charity_index, num_recursions = num_recursions + 1, too_late_signups_list = too_late_signups_list)
            return True
        
    return False
    
# Assigns the participant to the charity they were previously bumped from
def assign_to_previously_bumped_from_charity(participant_id, df_signups, df_charities, index_to_skip, num_recursions, too_late_signups_list):
    num_participants_in_group = df_signups.loc[participant_id, 'num_participants']
    
    # Bump off the most recent addition to this group and force them to go to a different charity lol
    most_recent_addition_id = pop_most_recently_added_participant(index_to_skip, df_signups, df_charities)

    prepend_participant_to_charity(participant_id, index_to_skip, df_signups, df_charities) # Prepend the ID (to make infinite loops only occur after checking the whole list)

    # print ("ID " + str(participant_id) + " assigned to " + df_charities.loc[index_to_skip, 'charity'])

    # Assign the popped id to a charity (if possible, not this one which they got popped from)
    assign_to_charity(most_recent_addition_id, df_signups, df_charities, index_to_skip = index_to_skip, num_recursions = num_recursions + 1, too_late_signups_list = too_late_signups_list)
    return


def add_participant_to_charity(participant_id, charity_index, df_signups, df_charities):
    df_charities.at[charity_index, 'participant_count'] = df_charities.loc[charity_index, 'participant_count'] + df_signups.loc[participant_id, 'num_participants']
    df_charities.at[charity_index, 'participant_ids'].append(participant_id)


def prepend_participant_to_charity(participant_id, charity_index, df_signups, df_charities):
    df_charities.at[charity_index, 'participant_count'] = df_charities.loc[charity_index, 'participant_count'] + df_signups.loc[participant_id, 'num_participants']
    df_charities.at[charity_index, 'participant_ids'].insert(0, participant_id) 

def pop_most_recently_added_participant(charity_index, df_signups, df_charities):
    most_recent_addition_id = df_charities.at[charity_index, 'participant_ids'].pop()
    num_participants_for_that_id = df_signups.loc[most_recent_addition_id, 'num_participants']
    df_charities.at[charity_index, 'participant_count'] = df_charities.loc[charity_index, 'participant_count'] - num_participants_for_that_id
    return most_recent_addition_id

# HELPER FUNCTIONS ------------------------------------------------------------------------

# This code was originally written for placing volunteers in charities, and the column names in the dataframe reflect this.
# I can't be bothered renaming everything in the code, so I'm just gonna rename the columns here LOL
def rename_columns_and_fill_empty_ones(df_charities, df_signups):
    if 'further_details' not in df_signups.columns:
        df_signups['further_details'] = ["nan" for i in range(len(df_signups))]

    df_signups.rename(columns = {'further_details': 'further_club_details', 'names': 'volunteer_names', 'preferenced_options': 'chosen_charities'}, inplace = True)
    df_charities.rename(columns = {'place': 'charity'}, inplace = True)


# Creates a list of the column names for emails, so it can be edited for different group sizes. For 3, returns ['email_1', 'email_2', 'email_3']
def get_email_col_names():
    return ['email_' + str(x + 1) for x in range(MAX_NUM_EMAILS)]

# Takes the lists of ids which are assigned to each charity, and assigns them to charities. Does this in-place
def transform_participant_ids_into_emails(df_charities, df_signups):
    # Transform the IDs into email lists
    df_charities['participant_emails'] = [[] for i in range(len(df_charities))]
    for i in range(len(df_charities)):
        for signup_id in df_charities.loc[i, 'participant_ids']:

            if (df_charities.loc[i, 'charity'] not in df_signups.loc[signup_id, 'charity_list']): # To check my code hasn't broken
                print("The code has broken somewhere lol")
                print("Misassiged placement: ID " + str(signup_id) + " with place " + df_charities.loc[i, 'charity'])

            df_charities.at[i, 'participant_emails'] = df_charities.at[i, 'participant_emails'] + df_signups.at[signup_id, 'participant_emails']

    # Formatting the emails of participants in a way that's easy to just copy paste into outlook
    df_charities['participant_emails_as_string'] = ''
    for i in range(len(df_charities)):
        df_charities.at[i, 'participant_emails_as_string'] = "; ".join(df_charities.at[i, 'participant_emails'])


# Creating a dataframe with the info of everyone the algorithm couldn't place
def generate_dataframe_of_unassignable_people(too_late_signups_list, ids_with_no_existing_options, df_signups):
    all_unassignable_people = ids_with_no_existing_options + too_late_signups_list

    df_unable_to_be_placed = pd.DataFrame()
    df_unable_to_be_placed['internal_code_id'] = all_unassignable_people
    df_unable_to_be_placed['emails'] = [[] for i in range(len(all_unassignable_people))]
    df_unable_to_be_placed['picked_options'] = [[] for i in range(len(all_unassignable_people))]
    for i in range(len(df_unable_to_be_placed)):
        df_unable_to_be_placed.at[i, 'emails'] = df_signups.at[all_unassignable_people[i], 'participant_emails']
        df_unable_to_be_placed.at[i, 'picked_options'] = df_signups.at[all_unassignable_people[i], 'charity_list']
    df_unable_to_be_placed = df_unable_to_be_placed.loc[:, ['emails', 'picked_options']]

    reason = ["None of the options this group chose were valid (ie. existed in the list of charities)" for i in range(len(ids_with_no_existing_options))]
    reason = reason + ["All of the options this group chose were full, and they were the most recent signup and so get lowest priority" for i in range(len(too_late_signups_list))]
    df_unable_to_be_placed['reason'] = reason

    return df_unable_to_be_placed

# A function to get the list of emails of people who wanted further details for each of the given charity clubs
def get_further_details_email_lists(df_signups):

    # Initialise a defaultdict for empty lists
    emails_lists = defaultdict(list)

    # For each signup in the list, split their list of clubs to find more info about, then add them to the list of emails for that club
    for i in range(len(df_signups)):
        emails_lists['All emails'] = emails_lists['All emails'] + df_signups.at[i, 'participant_emails']
        clubs = str(df_signups.at[i, 'further_club_details']).split(',')
        for club in clubs:
            club = club.strip()
            emails_lists[club] = emails_lists[club] + df_signups.at[i, 'participant_emails']

    # Remove the signups who didn't want any further information
    if "nan" in emails_lists.keys():
        emails_lists.pop("nan")

    # Transform the lists into a string of emails which are easy to just copy paste
    for key in emails_lists.keys():
        emails_lists[key] = '; '.join(emails_lists[key])

    return pd.DataFrame(sorted(emails_lists.items()))


# Save the outputs, or return an easier-to-read error message if we can't open the file
def save_dataframes_as_excel(df_charities, df_unable_to_be_placed, df_repeated_emails, df_emails_for_each_club_further_info):
    df_charities_output = df_charities.loc[:, ['charity', 'min_number', 'max_number', 'participant_count', 'participant_emails_as_string']]
    df_charities_output.rename(columns = {'charity': 'place'},inplace = True)

    try:
        with pd.ExcelWriter(PATH.joinpath("placements_output.xlsx")) as writer:
            df_repeated_emails.to_excel(writer, "Repeated Emails")
            df_unable_to_be_placed.to_excel(writer, "Unassignable People")
            df_charities_output.to_excel(writer, "Places & Assigned Emails")
            df_emails_for_each_club_further_info.to_excel(writer, "Addresses for Further Details")
    except:
        print("ERROR")
        print("The program successfully ran, but couldn't save the output. This is probably because you still have the excel file open. Please try again")

# RUNNING CODE ----------------------------------------------------------------------------

def run_generator():
    # Read the dataframes
    df_charities = pd.read_excel(PLACES_FILEPATH) # TODO: Change to df_places for extensibility # TODO Amend so this is more extensible and expects the name "place", or just rename this column at start lols
    df_signups = pd.read_excel(SIGNUPS_FILEPATH)

    rename_columns_and_fill_empty_ones(df_charities, df_signups)

    # Before sorting the signups, check for duplicated emails and note any repeated emails
    df_repeated_emails = check_repeated_emails(df_signups)
    remove_commas(df_charities, df_signups)
    count_num_participants_in_group(df_signups)
    count_charities_picked_by_group(df_signups)
    get_charity_ratios(df_charities, df_signups)

    # Sort our dataframes so we vaguely optimally pick who to add to each charity
    sort_charities_by_ratio(df_charities, df_signups)
    sort_participants_by_group_size_and_num_charities_selected(df_signups)
    
    # Go through and assign all our participants, and record the people the algorithm couldn't place
    df_unassignable_people = assign_participants_to_charities(df_charities, df_signups)

    # Go through the participant ids assigned to charities, and transform them into human-readable emails. Also checks that all the assignments are valid
    transform_participant_ids_into_emails(df_charities, df_signups)

    # Get the emails of everyone who wanted further information about one of the options
    df_emails_for_each_club_further_info = get_further_details_email_lists(df_signups)

    # Save the various dataframes
    save_dataframes_as_excel(df_charities, df_unassignable_people, df_repeated_emails, df_emails_for_each_club_further_info)

    print("Program completed. You can now open the output excel. Please read output and adjust/rerun the code where necessary, in particular for unassignable people and duplicate emails")
    input("Press enter to close program")

# Something to keep the code open in case of errors
if __name__ == '__main__':
    try:
        run_generator() # The actual main function
    except BaseException:
        print("There has been an error. Please inform a nerd so they can figure out what it is and how to fix it lol\n")
        import sys
        print(sys.exc_info()[0])
        import traceback
        print(traceback.format_exc())
        print("Press Enter to close (after you've found a nerd and shown them this error message so they can fix it)")
        input() 











    

    

    

