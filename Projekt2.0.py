# -*- coding: utf-8 -*-
import xlrd
import re
from scipy import stats
from prettytable import PrettyTable


print "Opening the data set..."
wb = xlrd.open_workbook('FinalLogRAW.xlsx')
print "Done.", "\n"

print "Opening the evaluation file..."
wb2 = xlrd.open_workbook('spss.xlsx')
print "Done.", "\n"

#Settings
users = {}
sequence_window = 5
sequence_gap = 1
s_treshhold = 0.5
p_treshold = 0.05
possible_actions = []
print_ignore = False
print_unhandled = False
print_handled = False
print_warning = False



class User:
    def __init__(self, numberv):
        self.gender = None #1 = flicka, 0 = pojke
        self.grade = None # 1 = dålig, 2 = bra!
        self.maxlevel = None
        self.usernumber = numberv
        self.actions = []
        self.current_setting = None
        self.visited_settings = {}
        self.current_room = None
        self.visited_rooms = {'unknown': []}
        self.interacted_persons = []
        self.dialog = None
        self.artifacts = 0 
        self.activity = None
        self.level = None
        self.confirmed = False
        self.learned_correct = None
        self.learned_incorrect = None
        self.doing_test = False
        self.test_activity = None
        self.test_score = 0
        self.interact_person = None
        self.CTA_suggest = None
        self.user_attitude = None #attityden mot CTA_suggestion (reject/accept)
        self.learned_facts = []
        self.minute_spent = 0

    def set_setting(self, setting):
        if "Tidslottet" not in setting and setting in self.visited_settings.keys():
            self.add_action('TidsresaRE') #Tidsresa till en redan besökt plats!
        elif "Tidslottet" not in setting:
            self.add_action('Tidsresa') #Tidsresa till ny plats

        
        self.current_setting = setting
        if setting not in self.visited_settings:
            self.visited_settings[setting] = {}


    def set_room(self, room):
        self.current_room = room
        if room not in self.visited_rooms.keys():
            self.visited_rooms[room] = []

    def set_dialoge(self, dialoge):
        self.dialog = dialoge

    def set_artifacts(self, artifacts):
        self.artifacts = artifacts
        
    def set_activity(self, activity):
        self.activity = activity
        
    def set_level(self, level):
        self.level = level
        
    def set_confirmed(self, confirmed):
        if self.activity == "conceptsMap":
            self.add_action("LearnCM")
        elif self.activity == "timeLine":
            self.add_action("LearnTL")
        self.confirmed = confirmed
        
    def set_learned_correct(self, correct):
        self.learned_correct = correct
        
    def set_learned_incorrect(self, incorrect):
        self.learned_incorrect = incorrect
        
    def set_doing_test(self, test):
        self.doing_test = test
        
    def set_test_score(self, score):
        self.test_score = score
        
    def set_interact_person(self, person):
        try:
            if "Tidsslottet" not in self.current_setting and "Tidsalv" not in person[0:7] and "Chronos" not in person[0:7]:
                if person not in self.interacted_persons:
                    self.add_action("InteractPerson")
                else:
                    self.add_action("InteractPersonRE")
        except:
            if print_warning:
                print "Warning (Clicked_interactable_person: current_setting not set)"
        if person not in self.interacted_persons:
            self.interacted_persons.append(person)
        self.interact_person = person
        
    def set_CTA_suggest(self, suggest):
        self.CTA_suggest = suggest
        
    def set_user_attitude(self, attitude):
        self.user_attitude = attitude

    def set_new_learned_fact(self, fact):
        if fact not in self.learned_facts:
            self.learned_facts.append(fact)

    def add_interacted_artifact(self, artifact):
        try:
            a = self.visited_rooms[self.current_room] #Crasha innan man lägger till i actionlist.
            if "Tidsmaskin" not in artifact and "Tidslottet" not in self.current_setting and artifact not in self.visited_rooms[self.current_room]:
                self.add_action("Artifact") #Lägg ny artefakt!
            elif "Tidsmaskin" not in artifact and "Tidslottet" not in self.current_setting:
                self.add_action("ArtifactRE") #Lägg till redan klickad artefakt!

            if artifact not in self.visited_rooms[self.current_room]:
                self.visited_rooms[self.current_room].append(artifact)
        except:
            if print_warning:
                print "Warning (Clicked_interactable_artifact: Current_room not set)"
            if artifact not in self.visited_rooms['unknown']:
                self.visited_rooms['unknown'].append(artifact)

    def set_minigame(self, minigame):
        self.add_action("Lek")

    def set_grade(self, grade):
        self.grade = grade

    def set_maxlevel(self, level):
        self.maxlevel = level
        
    def add_action(self, action):
        self.actions.append(action)
        if action not in possible_actions:
            possible_actions.append(action)
        

def create_user(number):
    new_user = User(number)
    users[number] = new_user
    return new_user

def action_abstraction(ignore, ignoreable_tuples):
    print "Performing action abstraction..."
    current_user = 0
    for s in wb.sheets():
        row_amount = s.nrows
        for row in range(row_amount):
            if row > 0:
                try:
                    user_number = int(s.cell(row,1).value)
                except:
                    if int(s.cell(row, 0).value) == 231798:
                        print "End of document.", "\n"
                        break
                    print "failade med user_number", s.cell(row,1).value
                if user_number not in users.keys():
                    current_user = create_user(user_number)

                log_event = int(s.cell(row,0).value)
                event_type = s.cell(row,2).value
                context = s.cell(row, 3).value
                key = s.cell(row, 6).value
                value = s.cell(row, 7).value

                if user_number != current_user.usernumber:
                    current_user = users[user_number]

                if key in ignore or (key,value) in ignoreable_tuples:
                    #Ignorera allt detta
                    print_it(log_event, key, value, ignorev=print_ignore)
                    
                elif key == "EnteredSetting":
                    #Går in i ett nytt område
                    current_user.set_setting(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "StartingInRoom" or key == 'EnteredRoom':
                    #Går in i ett nytt rum
                    current_user.set_room(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "TriggeredDialog":
                    #Har en dialog med någon
                    current_user.set_dialoge(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "FinishedDialog":
                    #Dialogen slutar
                    current_user.set_dialoge(None)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "ActivitySelected":
                    #Väljer en aktivitet
                    current_user.set_activity(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "LevelSelected":
                    #Väljer svårighetsgrad
                    current_user.set_level(value) 
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "clickedConfirmLevel":
                    #Startar aktiviteten med vald svårighetsgrad
                    current_user.set_confirmed(True)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "LearnCorrectCM":
                    #Antal inlärda rätt i den kognitiva modellen
                    current_user.set_learned_correct(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "LearnIncorrectCM":
                    #Antal inlärda fel i den kognitiva modellen
                    current_user.set_learned_incorrect(value) 
                    print_it(log_event, key, value, handled=print_handled)
                elif key == "LearnCorrectTL":
                    #Antal inlärda rätt i timeline
                    print_it(log_event, key, value, handled=print_handled)
                    pass
                elif key == "LearnIncorrectTL":
                    #Antal inlärda fel i timeline
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "ClickedInteractiveArtifact" and value == u"Gör test in Slottet_Kontor":
                    #Eleven väljer att göra ett test
                    current_user.set_doing_test(True) 
                    print_it(log_event, key, value, handled=print_handled)

                elif key == "ClickedInteractiveArtifact" and value == "Othellospel in Slottet_Vardagsrum":
                    #Användaren spelar spel med sin agent
                    print_it(log_event, key, value, handled=print_handled)
                    pass
                elif key == "ClickedInteractiveArtifact" and value == "tidsaxelKnapp in Slottet_skolsal":
                    #Användaren använder tidslinjen
                    print_it(log_event, key, value, handled=print_handled)
                    pass
                    
                elif current_user.doing_test and key == "ActivityStarted":
                    #Visar vilket test eleven gör samt svårighetsgrad
                    string = value.split(';')
                    current_user.set_activity(string[0])
                    current_user.set_level(string[1]) 
                    print_it(log_event, key, value, handled=print_handled)
                                                   
                elif current_user.doing_test and "Test" in key:
                    #testresultaten aktiviteten
                    current_user.set_test_score(value) #använda?
                    current_user.set_doing_test(False)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "ClickedPositionedPerson":
                    #interagerar med en viss person
                    current_user.set_interact_person(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "CTA level":
                    if value == "SuggestedHigherLevel":
                        current_user.set_CTA_suggest("Higher")
                        print_it(log_event, key, value, handled=print_handled)
                    elif value == "RejectedLevelChallenge":
                        current_user.set_user_attitude("Reject")
                        print_it(log_event, key, value, handled=print_handled)
                    elif value == "AcceptedLevelChallenge":
                        current_user.set_user_attitude("Accept")
                        print_it(log_event, key, value, handled=print_handled)
                elif key == "LearnFact":
                    #om eleven (?) lärt sig ett fact
                    current_user.set_new_learned_fact(value)
                    print_it(log_event, key, value, handled=print_handled)
                    
                elif key == "ReturnToCastle" and value == "Confirmed":
                    #användaren åkte tillbaka från sin tidsresa (kanske??)
                    print_it(log_event, key, value, handled=print_handled)
                    pass
                elif key[0:6] == "Level:":
                    #här händer något??
                    print_it(log_event, key, value, handled=print_handled)
                elif key == "MiniGameEnabled":
                    # value = spelet som spelas (för skojs skull??)
                    current_user.set_minigame(value)
                    print_it(log_event, key, value, handled=print_handled)
                elif key == "MiniGameExited":
                    # value = spelet som avslutas (för skojs skull??)
                    print_it(log_event, key, value, handled=print_handled)
                elif key == "ClickedInteractiveArtifact":
                    #Klickat på en interaktiv artefak
                    current_user.add_interacted_artifact(value)
                    print_it(log_event, key, value, handled=print_handled)
                elif key == "GlobalLevelUp":
                    #Användaren har låst upp en ny nivå. value example = NewLevel;10
                    print_it(log_event, key, value, handled=print_handled)
                    pass
                else:
                    print_it(log_event, key, value, unhandled=print_unhandled)
    print "Done.\n"


def evaluate_users():
    print "Collecting the student evaluations..."
    for s in wb2.sheets():
        row_amount = s.nrows
        for row in range(row_amount):
            if row > 0:
                user_number = int(s.cell(row,1).value)
                grade = int(s.cell(row,5).value)
                maxlevel = int(s.cell(row,8).value)
                minutes = int(s.cell(row,6).value)
                gender = int(s.cell(row,4).value)

                users[user_number].set_grade(grade)
                users[user_number].set_maxlevel(maxlevel)
                users[user_number].minute_spent = minutes
                users[user_number].gender = gender
    print "Evaluation done.\n"
    
def sort_users(user_list, criteria):
    a = []
    a_girl = 0
    b = []
    b_girl = 0
    for key in user_list:
        if user_list[key].grade == criteria[0]:
            a.append(user_list[key])
        elif user_list[key].grade == criteria[1]:
            b.append(user_list[key])
    for user in a:
        a_girl += user.gender
    for user in b:
        b_girl += user.gender
    return (a, b)

def print_it(log_event, key, value, ignorev=False, unhandled=False, handled=False):    
    if ignorev:
        print "\t","\t","\t", log_event, key, ":", value                     
    elif unhandled:
        print "Unhandled event:", log_event, key, ":", value
    elif handled:
        print "\t",log_event, key, ":", value


def find_seqs(seq, gap):
    collect = []
    queue = [seq]

    for element in queue:
        if len(element) == 1:
            break
        if element not in collect:
            collect.append(element)
        for index in xrange(len(element)):
            temp = []
            for inde in xrange(len(element)):
                if index != inde:
                    temp.append(element[inde])
            queue.append(temp)
    return collect
            
def find_sequences(seq, max_seq, gap, collect, checked_seq):
    length = len(seq)
    index = 0
    
    while index+max_seq+gap <= length:
        sample = seq[index:max_seq+gap+index]
        checked_seq.append(sample)
        new_sequences = find_seqs(sample, gap)
        for sequence in new_sequences:
            if len(sequence) <= max_seq and sequence not in collect:
                collect.append(sequence)
        index += 1
    return (collect, checked_seq)

def make_string(list):
    string = ""
    for index in list:
        string += index
        string += ";"
    return string

def create_regex(pattern, regex_list):
    string = ""
    for element in pattern:
        if string == "":
            string += element
            string += ";"
        else:
            string += regex_list
            string += ";"
            string += element
            string += ";"
    return string
        
def sequence_in_user(seq, pattern, gap, regex_list):
    seq_string = make_string(seq)
    pattern = re.compile(create_regex(pattern, regex_list))
    match = re.search(pattern, seq_string)
    try:
        x = match.group()
        return True
    except:
        return False

def get_i_support(seq, pattern, gap, regex_list):
    seq_string = make_string(seq)
    pattern = re.compile(create_regex(pattern, regex_list))
    match = re.findall(pattern, seq_string)
    try:
        return len(match)
    except:
        return 0

def create_regex_list():
    regex_list = ""
    for action in possible_actions:
        if regex_list == "":
            regex_list += "("+action
        else:
            regex_list += "|"+action
    regex_list += ")?"
    return regex_list

def get_i(value_1, members_1, value_2=0, members_2=1):
    return (float(value_1)/float(members_1))-(float(value_2)/float(members_2))

def print_results(result):
    for key in result:
        new_list = []
        if key != "stats":
            for sequence in result[key]:
                g1 = 0
                g2 = 0
                for number in sequence[1]:
                    g1 += number
                for number in sequence[2]:
                    g2 += number
                new_list.append([sequence[0], g1, g2, get_i(g1, result["stats"][0], g2, result["stats"][1]), sequence[4]])
        new_list = sorted(new_list, key=lambda sequence: sequence[3])
        tablehead = ["Group: " + str(key), "Sequence", "Frequence g1", "Frequence g2", "i-support (g1-g2)", "p-value"]
        table = PrettyTable(tablehead)
        print_it = False
        for res in new_list:
            print_it = True
            table.add_row([''] +res)
        if print_it: print table

def start():
    CM_related = ['AgentHadNoPropositionCM', 'AgentAcceptsCorrectPropositionCM',
                  'UserProposeCorrectCM', 'CTARejectCorrectCM', 'AgentContradictCorrectPropositionCM',
                  'UserAffirmedCorrectPropositionCM', 'UserProposeIncorrectCM',
                  'AgentAcceptsIncorrectPropositionCM', 'AgentProposedCorrectCM',
                  'AcceptedAgentsCorrectPropositionCM', 'AgentUnlearnsIncorrectFactCM',
                  'RejectedAgentsIncorrectPropositionCM', 'AgentProposedIncorrectCM',
                  'AcceptedAgentsIncorrectPropositionCM','CTARejectIncorrectCM',
                  'AgentContradictIncorrectPropositionCM', 'UserAffirmedIncorrectPropositionCM',
                  'RejectedAgentsCorrectPropositionCM', 'AgentUnlearnsCorrectFactCM',
                  'UserWithdrewCorrectPropositionCM', 'UserWithdrewIncorrectPropositionCM'
                  ]
    
    TL_related = ['AgentLearnsNewCorrectFactTL', 'UserProposeCorrectTL', 'CTAintroErrTL',
                  'AgentProposedIncorrectTL','RejectedAgentsCorrectPropositionTL',
                  'CTARejectCorrectTL', 'UserAffirmedCorrectPropositionTL', 'UserProposeIncorrectTL',
                  'AgentLearnsNewIncorrectFactTL', 'AgentLearnsNewIncorrectFactTL',
                  'AgentHadNoPropositionTL', 'undid', 'AcceptedAgentsIncorrectPropositionTL',
                  'AgentContradictIncorrectPropositionTL', 'UserAffirmedIncorrectPropositionTL',
                  'AgentUnlearnsIncorrectPropositionTL', 'AgentProposedCorrectTL',
                  'AcceptedAgentsCorrectPropositionTL', 'RejectedAgentsIncorrectPropositionTL',
                  'AgentContradictCorrectPropositionTL', 'AgentUnlearnsCorrectPropositionTL',
                  'AgentAcceptsCorrectPropositionTL', 'AgentAcceptsIncorrectPropositionTL',
                  'UserWithdrewCorrectPropositionTL', 'UserWithdrewIncorrectPropositionTL',
                  'AgentUnlearnsCorrectFactTL', 'AgentLearnsCorrectFactTL', 'AgentUnlearnsIncorrectFactTL',
                  'AgentLearnsIncorrectFactTL'
                  ]
    #undid = ångrade något i timelinen

    ignore = ['Login', 'DisplayedDialogLine', 'ClickedDialogResponse', 'ClickedDoorToRoom', 'RedirectedTo',
              'clickedLevel', 'InfoTextShown', 'InfoTextClosed', 'ClickedAgreementButton', 'InfoImageShown',
              'MagnifiedPlacedIconShown', 'MagnifiedIconShown', 'KnowledgeUpdatedOnServer']
    ignore += CM_related
    ignore += TL_related
    
    ignoreable_tuples = [('ClickedInteractiveArtifact', 'konceptkartaKnapp in Slottet_skolsal'),
                         ('ClickedInteractiveArtifact', 'Tidsmaskin in Tidsmaskinrum'),
                         ('ReturnToCastle', 'Cancelled')]

    action_abstraction(ignore, ignoreable_tuples)

    regex_list = create_regex_list()

    evaluate_users()
    sorted_users = sort_users(users, [1,2])
    
    collect = []
    checked_seq = []
    
    print "Gathering frequent action sequences that occur in more than " + str(s_treshhold*100) + "% of the time, for any of the groups..."
    for group in sorted_users:
        for user in group:
            (collect, checked_seq) = find_sequences(user.actions, sequence_window, sequence_gap, collect, checked_seq)

    frequent_patterns = dict()

    for pattern in collect:
        string = make_string(pattern)
        for group in sorted_users:
            counter = 0
            found = 0
            for user in group:
                if sequence_in_user(user.actions, pattern, sequence_gap, regex_list):
                    found += 1
                counter += 1
            if float(found)/float(counter) >= s_treshhold and string not in frequent_patterns.keys():
                frequent_patterns[string] = [pattern, [], [], None, None] #pattern, group1 i-values, group2 i-values, t-värde, p-värde
    print "Done. Found " + str(len(frequent_patterns)) + ".\n"


    print "Collecting i-support values..."
    for key in frequent_patterns:
        for user in sorted_users[0]:
            frequent_patterns[key][1].append(get_i_support(user.actions, frequent_patterns[key][0], sequence_gap, regex_list))
        for user in sorted_users[1]:
            frequent_patterns[key][2].append(get_i_support(user.actions, frequent_patterns[key][0], sequence_gap, regex_list))
    print "Done.\n"


    print "Calculating t-test..."
    for key in frequent_patterns:
        result = stats.ttest_ind(frequent_patterns[key][1], frequent_patterns[key][2])
        frequent_patterns[key][3] = result[0]
        frequent_patterns[key][4] = result[1]
    print "Done.\n"


    
    results ={"both":[], "only1":[], "only2":[]}


    for pattern in frequent_patterns:
        if frequent_patterns[pattern][1] == []:
            results["only2"].append(frequent_patterns[pattern])
        elif frequent_patterns[pattern][2] == []:
            results["only1"].append(frequent_patterns[pattern])
        elif frequent_patterns[pattern][4] <= p_treshold:
            results["both"].append(frequent_patterns[pattern])

    group_1_actions = 0
    group_1_playtime = 0
    group_1_members = 0
    group_1_girls = 0

    group_2_actions = 0
    group_2_playtime = 0
    group_2_members = 0
    group_2_girls = 0

    for user in sorted_users[0]:
        group_1_actions += len(user.actions)
        group_1_playtime += user.minute_spent
        group_1_members += 1
        group_1_girls += user.gender
    for user in sorted_users[1]:
        group_2_actions += len(user.actions)
        group_2_playtime += user.minute_spent
        group_2_members += 1
        group_2_girls += user.gender

    results["stats"] = [group_1_members, group_2_members]

        
    print "Group 1 statistics:\n Members: "+str(group_1_members)+"\n \t Girls: " + str(group_1_girls) + "\n \t Boys: " + str(group_1_members - group_1_girls) + "\n Total playtime (avg): "+str(group_1_playtime)+" ("+str(group_1_playtime/group_1_members)+")\n Total actions (avg): "+str(group_1_actions)+" ("+str(group_1_actions/group_1_members)+")\n"

    print "Group 2 statistics:\n Members: "+str(group_2_members)+"\n \t Girls: " + str(group_2_girls) + "\n \t Boys: " + str(group_2_members - group_2_girls) + "\n Total playtime (avg): "+str(group_2_playtime)+" ("+str(group_2_playtime/group_2_members)+")\n Total actions (avg): "+str(group_2_actions)+" ("+str(group_2_actions/group_2_members)+")\n"
            
    print "Found:\n" + str(len(results["both"])) + " statistically significant action sequences.\n", str(len(results["only1"])) + " that occured in only sample 1.\n", str(len(results["only2"])) + " that occured in only sample 2.\n"
    return results


result = start()
print_results(result)
