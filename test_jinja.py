from jinja2 import Environment,FileSystemLoader, select_autoescape



env = Environment(
    loader=FileSystemLoader("./"),
    autoescape=select_autoescape()
)
template = env.get_template("templ.html")
output=open("emploi.html","w")
output.write(
    template.render(
        a_variable="", 
        niveau="Niveau L2",
        emploi=[
            {"creneau":"08H00sss-09H00","dimanche":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","lundi":"G1', 'TypeSceance': 'TD', 'sceance': 'Comp', 'ProfName': 'Amrane', 'Salle': 'INF2","mardi":"G1', 'TypeSceance': 'Cours', 'sceance': 'PS', 'ProfName': 'Redjil', 'Salle': 'A11","mercredi":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","jeudi":"G1', 'TypeSceance': 'TP', 'sceance': 'SE', 'ProfName': 'Hariati', 'Salle': 'S1"},
            {"creneau":"09H15-10H15","dimanche":"G1', 'TypeSceance': 'TP', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'S1","lundi":"G1', 'TypeSceance': 'TD', 'sceance': 'PS', 'ProfName': 'Redjili', 'Salle': 'AG51","mardi":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","mercredi":"G1', 'TypeSceance': 'TD', 'sceance': 'Comp', 'ProfName': 'Sari', 'Salle': 'AG51","jeudi":"G1', 'TypeSceance': 'TP', 'sceance': 'GL', 'ProfName': 'Atil', 'Salle': 'S1"},
            ##{"creneau":"","dimanche":"Seance 2 du jour 1","lundi":"Seance 2 du jour 2","mardi":"Seance 2 du jour 3","mercredi":"Seance 2 du jour 4","jeudi":"Seance 2 du jour 5"},
            
        ]))
output.close()