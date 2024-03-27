def main():
    Student_List = ["Maria Hisham Elbuni","Belar Bilal Khader","Nadia Ali","Aiyah Ali","Sarah Ali","Miriam Ali","Reem Nor","Maryam Nasir Zekraoui","Nassira Nasir Zekraoui","Sereen Alder","Jury Amer Khader","Romeysa Loujan","Maryam Hisham Elbuni","Myra Eyad Ghannam","Nada Alder","Zayna GGG Khader","Rovan Ahmed","Rajen Ahmed","Ibtihal Mohomad Hasod","Marjan Uddin","Salma Mohamed Hussein","Rumeysa Khurshid Rosimatora","Elif Oral","Eslem Oral","Hind Mohomad Hasod","Meryem Oral","Shaymaa Moustafa","Jury Hatem","Judy Hatem","Rokaya Ali"]


    for i in Student_List:
        print(f'=FILTER(Student_List!A2:X,Student_List!C2:C="{i}")')

if __name__ == "__main__":
    main()