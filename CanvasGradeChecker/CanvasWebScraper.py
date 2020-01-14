from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
from time import sleep
import re
from UpdateFromCSV import UpdateFromCSV

driver = webdriver.Chrome(r"C:\Users\Reece\Desktop\Programming\Python\Storage\chromedriver.exe")
try:
    driver.get("https://canvas.vt.edu/")
    userElem = driver.find_element_by_name("j_username")
    userElem.send_keys("rpyankey")
    passElem = driver.find_element_by_name("j_password")
    passElem.send_keys("")

    # ask user to login
    while True:
        print("waiting for user to log in...")
        sleep(3)
        try:
            driver.find_element_by_class_name("ic-DashboardCard__link")
            print("logged in sucessfully")
            break
        except Exception:
            pass
    sleep(8)

    # find relevant class elements on dashboard
    class_link_elements = driver.find_elements_by_class_name("ic-DashboardCard__link")
    class_name_elements = driver.find_elements_by_class_name("ic-DashboardCard__header-subtitle")

    class_links = []  # class link/href
    class_names = []  # class name (ex: ENGE 1215)

    for elem in class_link_elements:
        class_links.append(elem.get_attribute("href"))

    for elem in class_name_elements:
        # get shortened course identifier (ex. ENGR 1045)
        result = re.search(r'[A-Z]{4}[\s_]?[0-9]{4}', elem.text)

        if result is None:
            currPos = len(class_names)
            class_links.pop(currPos)  # get rid of corresponding href
            continue

        # standardize result format and insert into class_names
        string = result.group(0)
        class_names.append(string[:4] + " " + string[-4:])

    print(class_links)
    print(class_names)

    for n in range(len(class_links)):
        class_link = class_links[n]
        class_name = class_names[n]
        print("gathering data for: "+class_name)

        # goto link
        driver.get(class_link + "/grades")
        html = driver.page_source
        soup = BeautifulSoup(html, features="html.parser")
        table = soup.find("table", id="grades_summary")

        # grab all table data
        rows = table.find_all("tr", class_="student_assignment")
        table_data = {"name": [], "date": [], "score": [], "max_score": [], "type": []}
        for row in rows:
            # skip elements that aren't actually assignments
            if "group_total" in row["class"] or "final_grade" in row["class"]:
                continue
            
            #add name of assignment
            table_data["name"].append(row.find("a").text)
            
            #add type of assignment
            table_data["type"].append(row.find("div", class_="context").text)
            
            #add due date of assignment
            date = row.find("td", class_="due").text
            formatted_date = re.search(r"[A-Za-z]{3}\s\d{1,2}", date)
            if formatted_date:
                table_data["date"].append(formatted_date.group(0))
            else:
                table_data["date"].append("N/A")
            
            #add grade of assinment
            score = row.find("span", class_="original_score").text
            formattedScore = re.search(r"\S+", score)
            if formattedScore:
                table_data["score"].append(formattedScore.group(0))
            else:
                table_data["score"].append("N/A")
            
            #add maximum score of assignment
            max_score = row.find("td", class_="points_possible").text
            formatted_max_score = re.search(r"\S+", max_score)
            table_data["max_score"].append(formatted_max_score.group(0))  # should be guaranteed to exist
            # TODO: add weight of types using <table class="summary">


        # store into csv
        table = pd.DataFrame(table_data)
        table.to_csv(class_name + ".csv")


    UpdateFromCSV(class_names)
finally:
    driver.quit()
