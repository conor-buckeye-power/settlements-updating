from scripts.ftr_lmp_updater import da_ftr_lmp_updater, rt_ftr_lmp_updater
from scripts.load_without_losses import update_load_without_losses
from scripts.get_miso_lmps import update_miso_lmps

print('========== Beginning Settlement Updates Automation ==========\n')

# Get user input for YEAR and MONTH to update, ensure that the input is an int otherwise fail
exit_script = False
try:
    year = int(input('Provide a year to update [yyyy]: '))
except ValueError:
    print('ERROR: Values inputted may not be integers')
    exit_script = True
try:
    month = int(input('Provide a month to update  [mm]: '))
except ValueError:
    print('ERROR: Values inputted may not be integers')
    exit_script = True

# Update MISO LMPs
update_miso_lmps(year, month, exit_script)

# Update FTR LMPs
da_ftr_lmp_updater(year, month, exit_script)
rt_ftr_lmp_updater(year, month, exit_script)

# Update Load without Losses
update_load_without_losses_check = input('\nUpdate load without losses? [y/n]: ')
if update_load_without_losses_check == 'y':
    update_load_without_losses(year, month, exit_script)

print('\n========== Settlement Automation Tasks Completed ==========')
