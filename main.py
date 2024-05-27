import pandas as pd
from itertools import combinations
from collections import defaultdict


file_path = input("Enter path of data file") #/Users/yuriihavryliuk/Desktop/университет/data_ready.xlsx"
df = pd.read_excel(file_path)


project_persons = defaultdict(set)
for _, row in df.iterrows():
    project = row['Project']
    person = row['Person']
    project_persons[project].add(person)

person_links = defaultdict(lambda: defaultdict(lambda: {'projects': [], 'weight': 0}))


for project, persons in project_persons.items():
    for person1, person2 in combinations(persons, 2):
        if person1 > person2:
            person1, person2 = person2, person1
        person_links[person1][person2]['projects'].append(project)
        person_links[person1][person2]['weight'] += 1


data = []
for person1, links in person_links.items():
    for person2, info in links.items():
        projects_str = '_'.join(info['projects'])
        weight = info['weight']
        data.append([person1, person2, projects_str, weight])


result_df = pd.DataFrame(data, columns=['Person1', 'Person2', 'Projects', 'Weight'])


with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    result_df.to_excel(writer, sheet_name='Person Links', index=False)

print("Завдання виконано успішно!")

