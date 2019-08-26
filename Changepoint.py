import win32com.client, datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

def Get_Calander_Data():
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")

    apps = ns.GetDefaultFolder(9).Items
    apps.Sort("[Start]")
    apps.IncludeRecurrences = "True"

    today = datetime.datetime.today()
    tomorrow = datetime.timedelta(days=1)+today
    end = tomorrow.date().strftime("%m/%d/%Y")
    four_days_before = datetime.timedelta(days=-4)+today
    start = four_days_before.date().strftime("%m/%d/%Y")

    apps = apps.Restrict("[Start] >= '" + start + "' AND [END] <= '" + end + "'")
    events={'Start':[],'Subject':[],'Duration':[]}

    for a in apps:
        # adate = datetime.datetime.fromtimestamp(int(a.Start))
        events['Start'].append(a.Start)
        events['Subject'].append(a.Subject)
        events['Duration'].append(a.Duration)

    workdays = [datetime.timedelta(),datetime.timedelta(),datetime.timedelta(),datetime.timedelta(),datetime.timedelta()]
    for a in range(len(events['Start'])):
        isTraining = str(events['Subject'][a]).upper().find('TRAINING')
        if(isTraining == -1):
            pass
        else:
            workdays[events['Start'][a].weekday()] += datetime.timedelta(minutes=events['Duration'][a])

        # print('Start:\t\t' + str(events['Start'][a]))
        # print('Day:\t\t' + str(events['Start'][a].weekday()))
        # print('Subject:\t' + str(events['Subject'][a]).upper())
        # print('Duration:\t' + str(datetime.timedelta(minutes=events['Duration'][a]))[:-3])

    return workdays



def Fill_Timesheet():
    URL = "http://changepoint.sedgwick.com:8080/"

    options = webdriver.ChromeOptions()
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option('detach', True)
    driver = webdriver.Chrome(options = options)
    driver.implicitly_wait(10)
    driver.get(URL)
    win_before = driver.window_handles[0]
    driver.implicitly_wait(10)
    win_after = driver.window_handles[1]
    driver.switch_to.window(win_after)

    driver.find_element_by_id('lnkMPERSONAL').click()
    driver.find_element_by_id('aTime').click()

    try:
        workday_hours = Get_Calander_Data()
        iframe = driver.find_elements_by_tag_name('iframe')[0]
        driver.switch_to.frame(iframe)
        time_entry = driver.find_element_by_id('tblTimeSheet_tblMain').find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
        for project in range(len(time_entry)):
            for i in range(4, 9):
                if(str(time_entry[project].find_elements_by_tag_name('td')[i].get_attribute("title")).upper() == 'Training'.upper()):
                    time_entry[project].find_elements_by_tag_name('td')[i].click()
                    time_entry[project].find_elements_by_tag_name('td')[i].find_element_by_tag_name('nobr').find_element_by_tag_name('input').send_keys(str(workday_hours[i-4])[0] + '.' + str(workday_hours[i-4])[2:4])
                else:
                    time_entry[project].find_elements_by_tag_name('td')[i].click()
                    time_entry[project].find_elements_by_tag_name('td')[i].find_element_by_tag_name('nobr').find_element_by_tag_name('input').send_keys(str(8.00 - float(str(workday_hours[i-4])[0] + '.' + str(workday_hours[i-4])[2:4])))
    except NoSuchElementException:
        print("Pin Projects to Time Sheet")

    # driver.find_element_by_id('Master_tdSubmit').click()


if __name__ == "__main__":
    Fill_Timesheet()
