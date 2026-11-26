"""Public text and navigation for the read-only introduction."""

CONTENT = {
    'B3': ('Retirement planner', 'report_title'),
    'B5': ('Use the fictional plan to see whether income and savings cover costs through age 95.', 'report_body'),
    'B7': ('Try the example', 'report_section'),
    'B8': ('In desktop Excel for Mac, keep Example plan selected on Inputs. Open Summary to find the base result for retirement at 60.', 'report_body'),
    'B10': ('Compare retirement at 65', 'report_section'),
    'B11': ('Read the saved market result first. Then choose Retire at 65 in Inputs B4. Summary updates the base plan and marks the saved risk result out of date.', 'report_body'),
    'B12': ('Open Inputs', 'report_link'),
    'B14': ('Run market risk or print', 'report_section'),
    'B15': ('Use 100 paired paths to explore or 1,000 for a larger sample. Follow your macro policy when you run market risk. Print the base plan with its risk status anytime.', 'report_body'),
    'B16': ('Open Summary', 'report_link'),
    'B18': ('Save a private copy before entering your figures', 'report_section'),
    'B19': ('Keep the .xlsm format. Use the sheet tabs if a link is unavailable. Appendix opens the detailed schedules.', 'report_body'),
    'B20': ('Open Appendix', 'report_link'),
}

LINKS = {'B12': 'Inputs!A1', 'B16': 'Summary!B3', 'B20': 'Appendix!B3'}
