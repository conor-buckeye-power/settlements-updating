from scripts.get_miso_lmps import update_miso_lmps

print(' ========== Beginning Settlement Updates Automation ========== \n')

# Update MISO LMPs
print('Updating Story County LMPs')
try:
    update_miso_lmps()
except:
    print('Failed to Update MISO LMPs (unknown error)\n')

print(' ========== Settlement Automation Tasks Completed ========== ')
