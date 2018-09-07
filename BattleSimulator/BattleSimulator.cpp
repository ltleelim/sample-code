#include <assert.h>
#include <math.h>
#include <stdlib.h>

#include <fstream>
#include <locale>
#include <string>

#include <Windows.h>

#include <XLCALL.H>

#include "ExcelCallbacks.h"
#include "VBACallbacks.h"

#include "EventQueue.h"

#include "BattleSimulator.h"


/* #pragma preprocessor linker directive to export functions */
#define EXPORT comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)


/* prefix non-negative numbers with a space to match VBA formatting */
#define SHOWSPACE(n) (((n) < 0) ? "" : " ") << (n)


#if LOG
#define OPENLOG(logFile, logBattles) \
        OpenLog(logFile, logBattles)
#define LOGSIMULATIONINFO(logFile, randomness, rngSeed, skipWeakerSpecialAttacks) \
        LogSimulationInfo(logFile, randomness, rngSeed, skipWeakerSpecialAttacks)
#define LOGPOKEMONINFO(logFile, role, pokedexNum, level, staminaIV, attackIV, defenseIV, transforms) \
        LogPokemonInfo(logFile, role, pokedexNum, level, staminaIV, attackIV, defenseIV, transforms)
#define LOGATTACKINFO(logFile, attackName, effectiveness, damage, energy, damageStart, duration) \
        LogAttackInfo(logFile, attackName, effectiveness, damage, energy, damageStart, duration)
#define LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, playerEvent) \
        LogEvent(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, playerEvent)
#define LOGNEWLINE(logFile) \
        LogNewline(logFile)
#define CLOSELOG(logFile) \
        CloseLog(logFile)
#else
#define OPENLOG(logFile, logBattles)
#define LOGSIMULATIONINFO(logFile, randomness, rngSeed, skipWeakerSpecialAttacks)
#define LOGPOKEMONINFO(logFile, role, pokedexNum, level, staminaIV, attackIV, defenseIV, transforms)
#define LOGATTACKINFO(logFile, attackName, effectiveness, damage, energy, damageStart, duration)
#define LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, playerEvent)
#define LOGNEWLINE(logFile)
#define CLOSELOG(logFile)
#endif


/* Excel Boolean is 2 bytes */
typedef __int16 ExcelBoolean;


inline int Min(int number1, int number2)
{
    return (number1 < number2) ? number1 : number2;
}


inline int Max(int number1, int number2)
{
    return (number1 > number2) ? number1 : number2;
}


int RandomInterval(int expectedInterval, int intervalRandomness)
{
    int intervalStart;

    intervalStart = expectedInterval - intervalRandomness / 2;
    return intervalStart + (int) ((intervalRandomness + 1) * (double) rand() / (RAND_MAX + 1));
}


#if 0
/* for reference */
std::string WStringToString(const std::wstring &wStr)
{
    std::wstring_convert<std::codecvt<wchar_t, char, std::mbstate_t>> converter;

    return converter.to_bytes(wStr);
}
#endif


#if LOG
std::string XLOPER12StrToString(const XLOPER12 &operand)
{
    std::wstring                                                      operandWStr;
    std::wstring_convert<std::codecvt<wchar_t, char, std::mbstate_t>> converter;

    assert(operand.xltype == xltypeStr);
    /* construct null terminated string */
    operandWStr.assign(&operand.val.str[1], operand.val.str[0]);
    /* convert wide string to string */
    return converter.to_bytes(operandWStr);
}
#endif


#if LOG
void OpenLog(std::ofstream &logFile, bool logBattles)
{
#if !THREADSAFE
    XLOPER12    workbookName, workbookPath, pathSeparator;
    std::string workbookNameStr, workbookPathStr, pathSeparatorStr;
#endif
    std::string logFileNameStr;

    if (logBattles && !logFile.is_open()) {
#if THREADSAFE
        logFileNameStr = "battle simulator log.txt";
#else
        workbookName = ActiveWorkbookName();
        workbookNameStr = XLOPER12StrToString(workbookName);
        workbookPath = ActiveWorkbookPath();
        workbookPathStr = XLOPER12StrToString(workbookPath);
        pathSeparator = PathSeparator();
        pathSeparatorStr = XLOPER12StrToString(pathSeparator);
        FREE(3, &workbookName, &workbookPath, &pathSeparator);
        logFileNameStr = workbookPathStr + pathSeparatorStr + workbookNameStr;
        logFileNameStr.replace(logFileNameStr.rfind(".xlsm"), 5, " log.txt");
#endif
        logFile.open(logFileNameStr, std::ios::out | std::ios::app);
        if (logFile.fail()) {
#if !THREADSAFE
            MsgBox(L"\030Opening log file failed.");
#endif
        }
    }
}
#endif


#if LOG
void LogSimulationInfo(std::ofstream &logFile, bool randomness, int rngSeed, bool skipWeakerSpecialAttacks)
{
    if (logFile.is_open()) {
        if (randomness) {
            logFile << "simulation uses random behavior\n";
            logFile << SHOWSPACE(rngSeed) << " random number generator seed\n";
        }
        else {
            logFile << "simulation uses expected behavior\n";
        }
        if (skipWeakerSpecialAttacks) {
            logFile << "simulation skips weaker special attacks\n";
        }
        else {
            logFile << "simulation always uses special attacks\n";
        }
    }
}
#endif


#if LOG
void LogPokemonInfo(std::ofstream &logFile, const std::string &role, int pokedexNum, double level, int staminaIV, int attackIV, int defenseIV, bool transforms)
{
    XLOPER12    name;
    std::string nameStr;

    if (logFile.is_open()) {
        logFile << role << ":\n";
        if (transforms) {
            logFile << "Ditto -> ";
        }
        name = VLookupString(pokedexNum, GetNamedRange(L"\017Species!Species"), 2, false);
        nameStr = XLOPER12StrToString(name);
        FREE(1, &name);
        logFile << nameStr << "\n";
        logFile << SHOWSPACE(level) << " level\n";
        logFile << SHOWSPACE(staminaIV) << " stamina IV\n";
        logFile << SHOWSPACE(attackIV) << " attack IV\n";
        logFile << SHOWSPACE(defenseIV) << " defense IV\n";
    }
}
#endif


#if LOG
void LogAttackInfo(std::ofstream &logFile, const XLOPER12 &attackName, double effectiveness, int damage, int energy, int damageStart, int duration)
{
    std::string attackNameStr;

    if (logFile.is_open()) {
        attackNameStr = XLOPER12StrToString(attackName);
        logFile << attackNameStr << "\n";
        logFile << SHOWSPACE(effectiveness) << " type effectiveness\n";
        logFile << SHOWSPACE(damage) << " damage\n";
        logFile << SHOWSPACE(energy) << " energy\n";
        logFile << SHOWSPACE(damageStart) << " damage start\n";
        logFile << SHOWSPACE(duration) << " duration\n";
    }
}
#endif


#if LOG
void LogEvent(std::ofstream &logFile, int battleTimer, int attackerBattleHP, int attackerEnergy, int defenderBattleHP, int defenderEnergy,
              const std::string &playerEvent)
{
    if (logFile.is_open()) {
        logFile << SHOWSPACE(battleTimer) << " "
                << SHOWSPACE(attackerBattleHP) << " " << SHOWSPACE(attackerEnergy) << " "
                << SHOWSPACE(defenderBattleHP) << " " << SHOWSPACE(defenderEnergy) << " "
                << playerEvent << "\n";
    }
}
#endif


#if LOG
void LogNewline(std::ofstream &logFile)
{
    if (logFile.is_open()) {
        logFile << std::endl;
    }
}
#endif


#if LOG
void CloseLog(std::ofstream &logFile)
{
    if (logFile.is_open()) {
        logFile.close();
    }
}
#endif


BOOL CALLBACK EnumWindowsProc(HWND hWnd, bool &calledFromExcelDialog)
{
    WCHAR classNameStr[sizeof "bosa_sdm_XL"];
    int   result;

    /* get only the number of characters being matched */
    result = GetClassName(hWnd, classNameStr, sizeof "bosa_sdm_XL");
    assert(result);
    /* class names are case insensitive */
    if (!_wcsicmp(classNameStr, L"bosa_sdm_XL")) {
        calledFromExcelDialog = true;
        /* stop enumerating windows */
        return FALSE;
    }
    /* continue enumerating windows */
    return TRUE;
}


bool CalledFromExcelDialog(void)
{
    bool calledFromExcelDialog;

    calledFromExcelDialog = false;
    (void) EnumWindows((WNDENUMPROC) EnumWindowsProc, (LPARAM) &calledFromExcelDialog);
    return calledFromExcelDialog;
}


LPXLOPER12 WINAPI xlAddInManagerInfo12(XLOPER12 &action)
{
#pragma EXPORT
    /* separate static variables for thread safety */
    static XLOPER12 dllName, valueError;

    if ((action.xltype == xltypeNum && action.val.num == 1.0) || (action.xltype == xltypeInt && action.val.w == 1)) {
        dllName.val.str = L"\020Battle Simulator";
        dllName.xltype = xltypeStr;
        return &dllName;
    } else {
        valueError.val.err = xlerrValue;
        valueError.xltype = xltypeErr;
        return &valueError;
    }
}


int WINAPI xlAutoOpen(void)
{
#pragma EXPORT
    XLOPER12 xllName;
    XLOPER12 functionName, typeText, argumentText, macroType, category, functionHelp, argumentHelp1, argumentHelp2, result;
    int      returnValue;

    /* get XLL path and name */
    returnValue = Excel12(xlGetName, &xllName, 0);
    if (returnValue != xlretSuccess) return 0;

    /* register functions with Excel */
    functionName.xltype = xltypeStr;
    functionName.val.str = L"\025SpecialAttackIsWeaker";
    typeText.xltype = xltypeStr;
#if THREADSAFE
    typeText.val.str = L"\004AJJ$";
#else
    typeText.val.str = L"\003AJJ";
#endif
    argumentText.xltype = xltypeStr;
    argumentText.val.str = L"\053attacker_move_set_num,defender_move_set_num";
    macroType.xltype = xltypeInt;
    macroType.val.w = 1;
    category.xltype = xltypeStr;
    category.val.str = L"\020Battle Simulator";
    functionHelp.xltype = xltypeStr;
    functionHelp.val.str = L"\152Calculates whether the attacker's special attack is weaker than its fast attack, and returns TRUE or FALSE";
    argumentHelp1.xltype = xltypeStr;
    argumentHelp1.val.str = L"\042is the attacker's move set number.";
    argumentHelp2.xltype = xltypeStr;
    argumentHelp2.val.str = L"\042is the defender's move set number.";
    returnValue = Excel12(xlfRegister, &result, 12, &xllName, &functionName, &typeText, &functionName, &argumentText, &macroType, &category, nullptr, nullptr,
                          &functionHelp, &argumentHelp1, &argumentHelp2);
    if (returnValue != xlretSuccess) return 0;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\006Battle";
    typeText.xltype = xltypeStr;
#if THREADSAFE
    typeText.val.str = L"\004BJJ$";
#else
    typeText.val.str = L"\003BJJ";
#endif
    argumentText.xltype = xltypeStr;
    argumentText.val.str = L"\053attacker_move_set_num,defender_move_set_num";
    macroType.xltype = xltypeInt;
    macroType.val.w = 1;
    category.xltype = xltypeStr;
    category.val.str = L"\020Battle Simulator";
    functionHelp.xltype = xltypeStr;
    functionHelp.val.str = L"\103Returns the probability of the attacker winning versus the defender";
    argumentHelp1.xltype = xltypeStr;
    argumentHelp1.val.str = L"\042is the attacker's move set number.";
    argumentHelp2.xltype = xltypeStr;
    argumentHelp2.val.str = L"\042is the defender's move set number.";
    returnValue = Excel12(xlfRegister, &result, 12, &xllName, &functionName, &typeText, &functionName, &argumentText, &macroType, &category, nullptr, nullptr,
                          &functionHelp, &argumentHelp1, &argumentHelp2);
    if (returnValue != xlretSuccess) return 0;

    functionName.xltype = xltypeStr;
    functionName.val.str = L"\026DefenderSpeciesAverage";
    typeText.xltype = xltypeStr;
#if THREADSAFE
    typeText.val.str = L"\004BJJ$";
#else
    typeText.val.str = L"\003BJJ";
#endif
    argumentText.xltype = xltypeStr;
    argumentText.val.str = L"\053attacker_move_set_num,defender_move_set_num";
    macroType.xltype = xltypeInt;
    macroType.val.w = 1;
    category.xltype = xltypeStr;
    category.val.str = L"\020Battle Simulator";
    functionHelp.xltype = xltypeStr;
    functionHelp.val.str = L"\144Returns the average of the matchups between the attacker and all move sets of the defender's species";
    argumentHelp1.xltype = xltypeStr;
    argumentHelp1.val.str = L"\042is the attacker's move set number.";
    argumentHelp2.xltype = xltypeStr;
    argumentHelp2.val.str = L"\042is the defender's move set number.";
    returnValue = Excel12(xlfRegister, &result, 12, &xllName, &functionName, &typeText, &functionName, &argumentText, &macroType, &category, nullptr, nullptr,
                          &functionHelp, &argumentHelp1, &argumentHelp2);
    if (returnValue != xlretSuccess) return 0;

    return 1;
}


double TypeEffectiveness(const XLOPER12 &attackerMoveType, const XLOPER12 &defenderType1, const XLOPER12 &defenderType2)
{
    XLOPER12 typeMatchupsRange;
    XLOPER12 attackingTypesRange, defendingTypesRange;
    int      rowNum, colNum;
    double   effectiveness;

    typeMatchupsRange = GetNamedRange(L"\034'Type Matchups'!TypeMatchups");
    attackingTypesRange = GetNamedRange(L"\036'Type Matchups'!AttackingTypes");
    defendingTypesRange = GetNamedRange(L"\036'Type Matchups'!DefendingTypes");
    rowNum = Match(attackerMoveType, attackingTypesRange, 0);
    colNum = Match(defenderType1, defendingTypesRange, 0);
    effectiveness = IndexNumber(typeMatchupsRange, rowNum, colNum);
    /* defenderType2.xltype can be xltypeNil */
    if (defenderType2.xltype == xltypeStr) {
        colNum = Match(defenderType2, defendingTypesRange, 0);
        effectiveness *= IndexNumber(typeMatchupsRange, rowNum, colNum);
    }
    FREE(3, &typeMatchupsRange, &attackingTypesRange, &defendingTypesRange);
    return effectiveness;
}


ExcelBoolean WINAPI SpecialAttackIsWeaker(long attackerMoveSetNum, long defenderMoveSetNum)
{
#pragma EXPORT
    double   attackerLevel, defenderLevel;
    int      attackerAttackIV, defenderDefenseIV;
    XLOPER12 levelsRange, speciesRange;
    int      attackerPokedexNum, defenderPokedexNum;
    double   attackerCPMultiplier, defenderCPMultiplier;
    double   attackerAttack, defenderDefense;
    XLOPER12 moveSetsRange;
    XLOPER12 attackerFastAttackType, attackerSpecialAttackType;
    int      attackerFastAttackPower, attackerSpecialAttackPower;
    int      attackerFastAttackDuration, attackerSpecialAttackDuration;
    double   attackerFastAttackSTAB, attackerSpecialAttackSTAB;
    XLOPER12 defenderType1, defenderType2;
    double   effectiveness;
    int      attackerFastAttackDamage, attackerSpecialAttackDamage;
    int      longPressDuration;
    double   attackerFastAttackDPS, attackerSpecialAttackDPS;

    /* do not execute from dialog box */
    if (CalledFromExcelDialog()) return FALSE;

    /* get global inputs */
    attackerLevel = GetNamedNumber(L"\024Inputs!AttackerLevel");
    attackerAttackIV = (int) GetNamedNumber(L"\027Inputs!AttackerAttackIV");

    defenderLevel = GetNamedNumber(L"\024Inputs!DefenderLevel");
    defenderDefenseIV = (int) GetNamedNumber(L"\030Inputs!DefenderDefenseIV");

    /* calculate stats */
    levelsRange = GetNamedRange(L"\015Levels!Levels");
    speciesRange = GetNamedRange(L"\017Species!Species");
    attackerPokedexNum = attackerMoveSetNum / 1000000;
    attackerCPMultiplier = VLookupNumber(attackerLevel, levelsRange, 2, false);
    attackerAttack = (VLookupNumber(attackerPokedexNum, speciesRange, 6, false) + attackerAttackIV) * attackerCPMultiplier;

    defenderPokedexNum = defenderMoveSetNum / 1000000;
    defenderCPMultiplier = VLookupNumber(defenderLevel, levelsRange, 2, false);
    defenderDefense = (VLookupNumber(defenderPokedexNum, speciesRange, 7, false) + defenderDefenseIV) * defenderCPMultiplier;

    /* get move data */
    moveSetsRange = GetNamedRange(L"\024'Move Sets'!MoveSets");
    attackerFastAttackType = VLookupString(attackerMoveSetNum, moveSetsRange, 8, false);
    attackerFastAttackPower = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 9, false);
    attackerFastAttackDuration = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 12, false);
    attackerFastAttackSTAB = VLookupNumber(attackerMoveSetNum, moveSetsRange, 13, false);
    attackerSpecialAttackType = VLookupString(attackerMoveSetNum, moveSetsRange, 16, false);
    attackerSpecialAttackPower = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 17, false);
    attackerSpecialAttackDuration = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 20, false);
    attackerSpecialAttackSTAB = VLookupNumber(attackerMoveSetNum, moveSetsRange, 21, false);

    /* calculate damage against opponent */
    defenderType1 = VLookupString(defenderPokedexNum, speciesRange, 3, false);
    defenderType2 = VLookupString(defenderPokedexNum, speciesRange, 4, false);

    effectiveness = TypeEffectiveness(attackerFastAttackType, defenderType1, defenderType2);
    attackerFastAttackDamage = (int) (0.5 * attackerAttack / defenderDefense * attackerFastAttackPower * attackerFastAttackSTAB * effectiveness) + 1;

    effectiveness = TypeEffectiveness(attackerSpecialAttackType, defenderType1, defenderType2);
    attackerSpecialAttackDamage = (int) (0.5 * attackerAttack / defenderDefense * attackerSpecialAttackPower * attackerSpecialAttackSTAB * effectiveness) + 1;

    /* get battle parameters */
    longPressDuration = (int) GetNamedNumber(L"\030Inputs!LongPressDuration");

    /* calculate damage per second */
    attackerFastAttackDPS = attackerFastAttackDamage / (attackerFastAttackDuration / 1000.0);
    attackerSpecialAttackDPS = attackerSpecialAttackDamage / ((longPressDuration + attackerSpecialAttackDuration) / 1000.0);

    FREE(7, &levelsRange, &speciesRange, &moveSetsRange, &attackerFastAttackType, &attackerSpecialAttackType, &defenderType1, &defenderType2);

    /* return whether special attack DPS is less than fast attack DPS */
    return attackerSpecialAttackDPS <= attackerFastAttackDPS;
}


double WINAPI Battle(long attackerMoveSetNum, long defenderMoveSetNum)
{
#pragma EXPORT
    bool          skipWeakerSpecialAttacks;
    bool          randomness;
    int           rngSeed;
    long          numTrials;
    bool          logBattles;
    std::ofstream logFile;
    double        attackerLevel, defenderLevel;
    int           attackerStaminaIV, attackerAttackIV, attackerDefenseIV;
    int           defenderStaminaIV, defenderAttackIV, defenderDefenseIV;
    XLOPER12      levelsRange, speciesRange;
    int           attackerPokedexNum, defenderPokedexNum;
    double        attackerCPMultiplier, defenderCPMultiplier;
    int           attackerHP, defenderHP;
    bool          attackerTransforms, defenderTransforms;
    long          transformMoveNum;
    XLOPER12      fastAttacksRange;
    int           transformPower, transformEnergy, transformDamageStart, transformDuration, transformDamage;
    double        attackerAttack, attackerDefense;
    double        defenderAttack, defenderDefense;
    XLOPER12      moveSetsRange;
    XLOPER12      attackerFastAttackName, attackerSpecialAttackName, defenderFastAttackName, defenderSpecialAttackName;
    XLOPER12      attackerFastAttackType, attackerSpecialAttackType, defenderFastAttackType, defenderSpecialAttackType;
    int           attackerFastAttackPower, attackerSpecialAttackPower, defenderFastAttackPower, defenderSpecialAttackPower;
    int           attackerFastAttackEnergy, attackerSpecialAttackEnergy, defenderFastAttackEnergy, defenderSpecialAttackEnergy;
    int           attackerFastAttackDamageStart, attackerSpecialAttackDamageStart, defenderFastAttackDamageStart, defenderSpecialAttackDamageStart;
    int           attackerFastAttackDuration, attackerSpecialAttackDuration, defenderFastAttackDuration, defenderSpecialAttackDuration;
    double        attackerFastAttackSTAB, attackerSpecialAttackSTAB, defenderFastAttackSTAB, defenderSpecialAttackSTAB;
    XLOPER12      attackerType1, attackerType2, defenderType1, defenderType2;
    double        effectiveness;
    int           attackerFastAttackDamage, attackerSpecialAttackDamage, defenderFastAttackDamage, defenderSpecialAttackDamage;
    double        defensiveHPMultiplier;
    int           maxAttackerEnergy, maxDefenderEnergy;
    double        energyPerDamage;
    int           battleDuration, longPressDuration;
    int           offensiveInitialInterval;
    int           numDefensiveInitialIntervals;
    int           *defensiveInitialIntervals;
    int           defensiveInterval, defensiveIntervalRandomness;
    int           numDefensiveSpecialAttackDeferrals;
    double        defensiveSpecialAttackProbability;
    double        attackerFastAttackDPS, attackerSpecialAttackDPS;
    EventQueue    attackerEventQueue, defenderEventQueue;
    long          numWins;
    int           defenderTime;
    int           battleTimer, nextTime;
    int           attackerBattleHP, defenderBattleHP;
    int           attackerEnergy, defenderEnergy;
    int           numDefensiveSpecialAttackOpportunities;
    PlayerEvents  playerEvent;
    bool          specialAttack;
    int           interval;
    long          i;
 
    /* do not execute from dialog box */
    if (CalledFromExcelDialog()) return 0.0;
   
    /* get simulation settings */
    skipWeakerSpecialAttacks = GetNamedBoolean(L"\037Inputs!SkipWeakerSpecialAttacks");
    randomness = GetNamedBoolean(L"\021Inputs!Randomness");
    rngSeed = (int) GetNamedNumber(L"\016Inputs!RNGSeed");
    assert(rngSeed > 0);
    numTrials = (long) GetNamedNumber(L"\032Inputs!NumMonteCarloTrials");
    assert(numTrials > 0);
    if (numTrials > 1 && !randomness) {
#if !THREADSAFE
        MsgBox(L"\053Monte Carlo simulations require randomness.");
#endif
        return -1.0;
    }

    /* if enabled, print log to file */
    logBattles = GetNamedBoolean(L"\021Inputs!LogBattles");
    OPENLOG(logFile, logBattles);
    
    /* seed random number generator */
    srand(rngSeed);
    LOGSIMULATIONINFO(logFile, randomness, rngSeed, skipWeakerSpecialAttacks);
    LOGNEWLINE(logFile);
    
    /* get global inputs */
    attackerLevel = GetNamedNumber(L"\024Inputs!AttackerLevel");
    attackerStaminaIV = (int) GetNamedNumber(L"\030Inputs!AttackerStaminaIV");
    attackerAttackIV = (int) GetNamedNumber(L"\027Inputs!AttackerAttackIV");
    attackerDefenseIV = (int) GetNamedNumber(L"\030Inputs!AttackerDefenseIV");

    defenderLevel = GetNamedNumber(L"\024Inputs!DefenderLevel");
    defenderStaminaIV = (int) GetNamedNumber(L"\030Inputs!DefenderStaminaIV");
    defenderAttackIV = (int) GetNamedNumber(L"\027Inputs!DefenderAttackIV");
    defenderDefenseIV = (int) GetNamedNumber(L"\030Inputs!DefenderDefenseIV");

    /* calculate stats */
    levelsRange = GetNamedRange(L"\015Levels!Levels");
    speciesRange = GetNamedRange(L"\017Species!Species");
    attackerPokedexNum = attackerMoveSetNum / 1000000;
    attackerCPMultiplier = VLookupNumber(attackerLevel, levelsRange, 2, false);
    attackerHP = Max((int) ((VLookupNumber(attackerPokedexNum, speciesRange, 5, false) + attackerStaminaIV) * attackerCPMultiplier), 10);

    defenderPokedexNum = defenderMoveSetNum / 1000000;
    defenderCPMultiplier = VLookupNumber(defenderLevel, levelsRange, 2, false);
    defenderHP = Max((int) ((VLookupNumber(defenderPokedexNum, speciesRange, 5, false) + defenderStaminaIV) * defenderCPMultiplier), 10);

    attackerTransforms = false;
    defenderTransforms = false;
    if ((attackerPokedexNum == dittoPokedexNum) != (defenderPokedexNum == dittoPokedexNum)) {
        /* perform Ditto transformations */
        if (attackerPokedexNum == dittoPokedexNum) {
            transformMoveNum = (attackerMoveSetNum % 1000000) / 1000;
            attackerMoveSetNum = defenderMoveSetNum;
            attackerPokedexNum = defenderPokedexNum;
            attackerTransforms = true;
        }
        if (defenderPokedexNum == dittoPokedexNum) {
            transformMoveNum = (defenderMoveSetNum % 1000000) / 1000;
            defenderMoveSetNum = attackerMoveSetNum;
            defenderPokedexNum = attackerPokedexNum;
            defenderTransforms = true;
        }
        fastAttacksRange = GetNamedRange(L"\032'Fast Attacks'!FastAttacks");
        transformPower = (int) VLookupNumber(transformMoveNum, fastAttacksRange, 4, false);
        transformEnergy = (int) VLookupNumber(transformMoveNum, fastAttacksRange, 5, false);
        transformDamageStart = (int) VLookupNumber(transformMoveNum, fastAttacksRange, 6, false);
        transformDuration = (int) VLookupNumber(transformMoveNum, fastAttacksRange, 7, false);
        assert(transformPower == 0);
        transformDamage = 1;
        FREE(1, &fastAttacksRange);
    }

    attackerAttack = (VLookupNumber(attackerPokedexNum, speciesRange, 6, false) + attackerAttackIV) * attackerCPMultiplier;
    attackerDefense = (VLookupNumber(attackerPokedexNum, speciesRange, 7, false) + attackerDefenseIV) * attackerCPMultiplier;

    defenderAttack = (VLookupNumber(defenderPokedexNum, speciesRange, 6, false) + defenderAttackIV) * defenderCPMultiplier;
    defenderDefense = (VLookupNumber(defenderPokedexNum, speciesRange, 7, false) + defenderDefenseIV) * defenderCPMultiplier;

    /* get move data */
    moveSetsRange = GetNamedRange(L"\024'Move Sets'!MoveSets");
    attackerFastAttackName = VLookupString(attackerMoveSetNum, moveSetsRange, 7, false);
    attackerFastAttackType = VLookupString(attackerMoveSetNum, moveSetsRange, 8, false);
    attackerFastAttackPower = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 9, false);
    attackerFastAttackEnergy = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 10, false);
    attackerFastAttackDamageStart = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 11, false);
    attackerFastAttackDuration = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 12, false);
    attackerFastAttackSTAB = VLookupNumber(attackerMoveSetNum, moveSetsRange, 13, false);
    attackerSpecialAttackName = VLookupString(attackerMoveSetNum, moveSetsRange, 15, false);
    attackerSpecialAttackType = VLookupString(attackerMoveSetNum, moveSetsRange, 16, false);
    attackerSpecialAttackPower = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 17, false);
    attackerSpecialAttackEnergy = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 18, false);
    attackerSpecialAttackDamageStart = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 19, false);
    attackerSpecialAttackDuration = (int) VLookupNumber(attackerMoveSetNum, moveSetsRange, 20, false);
    attackerSpecialAttackSTAB = VLookupNumber(attackerMoveSetNum, moveSetsRange, 21, false);

    defenderFastAttackName = VLookupString(defenderMoveSetNum, moveSetsRange, 7, false);
    defenderFastAttackType = VLookupString(defenderMoveSetNum, moveSetsRange, 8, false);
    defenderFastAttackPower = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 9, false);
    defenderFastAttackEnergy = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 10, false);
    defenderFastAttackDamageStart = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 11, false);
    defenderFastAttackDuration = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 12, false);
    defenderFastAttackSTAB = VLookupNumber(defenderMoveSetNum, moveSetsRange, 13, false);
    defenderSpecialAttackName = VLookupString(defenderMoveSetNum, moveSetsRange, 15, false);
    defenderSpecialAttackType = VLookupString(defenderMoveSetNum, moveSetsRange, 16, false);
    defenderSpecialAttackPower = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 17, false);
    defenderSpecialAttackEnergy = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 18, false);
    defenderSpecialAttackDamageStart = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 19, false);
    defenderSpecialAttackDuration = (int) VLookupNumber(defenderMoveSetNum, moveSetsRange, 20, false);
    defenderSpecialAttackSTAB = VLookupNumber(defenderMoveSetNum, moveSetsRange, 21, false);

    /* calculate damage against opponent */
    attackerType1 = VLookupString(attackerPokedexNum, speciesRange, 3, false);
    attackerType2 = VLookupString(attackerPokedexNum, speciesRange, 4, false);

    defenderType1 = VLookupString(defenderPokedexNum, speciesRange, 3, false);
    defenderType2 = VLookupString(defenderPokedexNum, speciesRange, 4, false);

    LOGPOKEMONINFO(logFile, "attacker", attackerPokedexNum, attackerLevel, attackerStaminaIV, attackerAttackIV, attackerDefenseIV, attackerTransforms);
    effectiveness = TypeEffectiveness(attackerFastAttackType, defenderType1, defenderType2);
    attackerFastAttackDamage = (int) (0.5 * attackerAttack / defenderDefense * attackerFastAttackPower * attackerFastAttackSTAB * effectiveness) + 1;
    LOGATTACKINFO(logFile, attackerFastAttackName, effectiveness, attackerFastAttackDamage, attackerFastAttackEnergy, attackerFastAttackDamageStart,
                  attackerFastAttackDuration);

    effectiveness = TypeEffectiveness(attackerSpecialAttackType, defenderType1, defenderType2);
    attackerSpecialAttackDamage = (int) (0.5 * attackerAttack / defenderDefense * attackerSpecialAttackPower * attackerSpecialAttackSTAB * effectiveness) + 1;
    LOGATTACKINFO(logFile, attackerSpecialAttackName, effectiveness, attackerSpecialAttackDamage, attackerSpecialAttackEnergy, attackerSpecialAttackDamageStart,
                  attackerSpecialAttackDuration);
    LOGNEWLINE(logFile);

    LOGPOKEMONINFO(logFile, "defender", defenderPokedexNum, defenderLevel, defenderStaminaIV, defenderAttackIV, defenderDefenseIV, defenderTransforms);
    effectiveness = TypeEffectiveness(defenderFastAttackType, attackerType1, attackerType2);
    defenderFastAttackDamage = (int) (0.5 * defenderAttack / attackerDefense * defenderFastAttackPower * defenderFastAttackSTAB * effectiveness) + 1;
    LOGATTACKINFO(logFile, defenderFastAttackName, effectiveness, defenderFastAttackDamage, defenderFastAttackEnergy, defenderFastAttackDamageStart,
                  defenderFastAttackDuration);

    effectiveness = TypeEffectiveness(defenderSpecialAttackType, attackerType1, attackerType2);
    defenderSpecialAttackDamage = (int) (0.5 * defenderAttack / attackerDefense * defenderSpecialAttackPower * defenderSpecialAttackSTAB * effectiveness) + 1;
    LOGATTACKINFO(logFile, defenderSpecialAttackName, effectiveness, defenderSpecialAttackDamage, defenderSpecialAttackEnergy, defenderSpecialAttackDamageStart,
                  defenderSpecialAttackDuration);
    LOGNEWLINE(logFile);

    /* get battle parameters */
    defensiveHPMultiplier = GetNamedNumber(L"\034Inputs!DefensiveHPMultiplier");
    maxAttackerEnergy = (int) GetNamedNumber(L"\031Inputs!MaxOffensiveEnergy");
    maxDefenderEnergy = (int) GetNamedNumber(L"\031Inputs!MaxDefensiveEnergy");
    energyPerDamage = GetNamedNumber(L"\026Inputs!EnergyPerHPLost");
    battleDuration = (int) GetNamedNumber(L"\025Inputs!BattleDuration");
    longPressDuration = (int) GetNamedNumber(L"\030Inputs!LongPressDuration");
    offensiveInitialInterval = (int) GetNamedNumber(L"\037Inputs!OffensiveInitialInterval");
    numDefensiveInitialIntervals = (int) GetNamedNumber(L"\043Inputs!NumDefensiveInitialIntervals");
    defensiveInitialIntervals = new int[numDefensiveInitialIntervals];
    defensiveInitialIntervals[0] = (int) GetNamedNumber(L"\044Inputs!DefensiveFirstInitialInterval");
    defensiveInitialIntervals[1] = (int) GetNamedNumber(L"\045Inputs!DefensiveSecondInitialInterval");
    defensiveInitialIntervals[2] = (int) GetNamedNumber(L"\044Inputs!DefensiveThirdInitialInterval");
    defensiveInterval = (int) GetNamedNumber(L"\030Inputs!DefensiveInterval");
    defensiveIntervalRandomness = (int) GetNamedNumber(L"\042Inputs!DefensiveIntervalRandomness");
    numDefensiveSpecialAttackDeferrals = (int) GetNamedNumber(L"\051Inputs!NumDefensiveSpecialAttackDeferrals");
    defensiveSpecialAttackProbability = GetNamedNumber(L"\050Inputs!DefensiveSpecialAttackProbability");

    /* calculate damage per second */
    attackerFastAttackDPS = attackerFastAttackDamage / (attackerFastAttackDuration / 1000.0);
    attackerSpecialAttackDPS = attackerSpecialAttackDamage / ((longPressDuration + attackerSpecialAttackDuration) / 1000.0);

    /* if enabled, skip weaker special attacks */
    if (skipWeakerSpecialAttacks && attackerSpecialAttackDPS <= attackerFastAttackDPS) {
        /* disable special attacks by making energy requirement unreachable */
        attackerSpecialAttackEnergy = -(maxAttackerEnergy + 1);
    }

    /* perform Monte Carlo trials */
    numWins = 0;
    for (i = 0; i < numTrials; ++i) {

        /* set up event queues */
        attackerEventQueue.Initialize(1);
        if (attackerTransforms) {
            attackerEventQueue.Add(offensiveInitialInterval, PlayerStartsTransform);
        } else {
            attackerEventQueue.Add(offensiveInitialInterval, PlayerStartsAttack);
        }
        defenderEventQueue.Initialize(numDefensiveInitialIntervals);
        defenderTime = defensiveInitialIntervals[0];
        if (defenderTransforms) {
            defenderEventQueue.Add(defenderTime, PlayerStartsTransform);
        } else {
            defenderEventQueue.Add(defenderTime, PlayerStartsInitialAttack);
        }
        defenderTime += defensiveInitialIntervals[1];
        defenderEventQueue.Add(defenderTime, PlayerStartsInitialAttack);
        if (randomness) {
            defenderTime += RandomInterval(defensiveInitialIntervals[2], defensiveIntervalRandomness);
        } else {
            defenderTime += defensiveInitialIntervals[2];
        }
        defenderEventQueue.Add(defenderTime, PlayerStartsAttack);

        /* simulate battle */
        battleTimer = battleDuration;
        attackerBattleHP = attackerHP;
        defenderBattleHP = (int) (defenderHP * defensiveHPMultiplier);
        attackerEnergy = 0;
        defenderEnergy = 0;
        numDefensiveSpecialAttackOpportunities = 0;
        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "battle starts");
        while (battleTimer > 0 && attackerBattleHP > 0 && defenderBattleHP > 0) {

            /* count down timers to next event */
            nextTime = Min(attackerEventQueue.Timer(), defenderEventQueue.Timer());
            attackerEventQueue.CountDown(nextTime);
            defenderEventQueue.CountDown(nextTime);
            battleTimer -= nextTime;

            /* check if time for next attacker event */
            if (attackerEventQueue.Timer() == 0) {
                playerEvent = attackerEventQueue.Pop();

                /* attacker finishes action */
                switch (playerEvent) {
                case PlayerFinishesLongPress:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker finishes long press");
                    break;
                case PlayerLandsFastAttack:
                    /* attacker lands fast attack damage */
                    defenderBattleHP -= attackerFastAttackDamage;
                    defenderEnergy = Min(defenderEnergy + (int) round(attackerFastAttackDamage * energyPerDamage + tolerance), maxDefenderEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker lands fast attack");
                    break;
                case PlayerLandsSpecialAttack:
                    /* attacker lands special attack damage */
                    defenderBattleHP -= attackerSpecialAttackDamage;
                    defenderEnergy = Min(defenderEnergy + (int) round(attackerSpecialAttackDamage * energyPerDamage + tolerance), maxDefenderEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker lands special attack");
                    break;
                case PlayerLandsTransform:
                    /* attacker lands transform damage */
                    defenderBattleHP -= transformDamage;
                    defenderEnergy = Min(defenderEnergy + (int) round(transformDamage * energyPerDamage + tolerance), maxDefenderEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker lands transform");
                    break;
                case PlayerFinishesFastAttack:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker finishes fast attack");
                    break;
                case PlayerFinishesSpecialAttack:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker finishes special attack");
                    break;
                case PlayerFinishesTransform:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker finishes transform");
                    break;
                case PlayerFinishesInitialFastAttack:
                case PlayerFinishesInitialSpecialAttack:
                    assert(false);
                    break;
                }

                /* attacker performs next action */
                switch (playerEvent) {
                case PlayerStartsInitialAttack:
                    assert(false);
                    break;
                case PlayerStartsTransform:
                    /* attacker starts transform */
                    attackerEnergy = Min(attackerEnergy + transformEnergy, maxAttackerEnergy);
                    attackerEventQueue.Add(transformDamageStart, PlayerLandsTransform);
                    attackerEventQueue.Add(transformDuration, PlayerFinishesTransform);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker starts transform");
                    break;
                case PlayerStartsAttack:
                case PlayerFinishesFastAttack:
                case PlayerFinishesSpecialAttack:
                case PlayerFinishesTransform:
                    /* attacker starts next attack */
                    if (attackerEnergy >= -attackerSpecialAttackEnergy) {
                        /* special attack */
                        attackerEventQueue.Add(longPressDuration, PlayerFinishesLongPress);
                        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker starts long press");
                    } else {
                        /* fast attack */
                        attackerEnergy = Min(attackerEnergy + attackerFastAttackEnergy, maxAttackerEnergy);
                        attackerEventQueue.Add(attackerFastAttackDamageStart, PlayerLandsFastAttack);
                        attackerEventQueue.Add(attackerFastAttackDuration, PlayerFinishesFastAttack);
                        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker starts fast attack");
                    }
                    break;
                case PlayerFinishesLongPress:
                    /* attacker continues special attack */
                    attackerEnergy = attackerEnergy + attackerSpecialAttackEnergy;
                    attackerEventQueue.Add(attackerSpecialAttackDamageStart, PlayerLandsSpecialAttack);
                    attackerEventQueue.Add(attackerSpecialAttackDuration, PlayerFinishesSpecialAttack);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "attacker starts special attack");
                    break;
                }
            }

            /* check if time for next defender event */
            if (defenderEventQueue.Timer() == 0) {
                playerEvent = defenderEventQueue.Pop();

                /* defender finishes action */
                switch (playerEvent) {
                case PlayerFinishesLongPress:
                    assert(false);
                    break;
                case PlayerLandsFastAttack:
                    /* defender lands fast attack damage */
                    attackerBattleHP -= defenderFastAttackDamage;
                    attackerEnergy = Min(attackerEnergy + (int) round(defenderFastAttackDamage * energyPerDamage + tolerance), maxAttackerEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender lands fast attack");
                    break;
                case PlayerLandsSpecialAttack:
                    /* defender lands special attack damage */
                    attackerBattleHP -= defenderSpecialAttackDamage;
                    attackerEnergy = Min(attackerEnergy + (int) round(defenderSpecialAttackDamage * energyPerDamage + tolerance), maxAttackerEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender lands special attack");
                    break;
                case PlayerLandsTransform:
                    /* defender lands transform damage */
                    attackerBattleHP -= transformDamage;
                    attackerEnergy = Min(attackerEnergy + (int) round(transformDamage * energyPerDamage + tolerance), maxAttackerEnergy);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender lands transform");
                    break;
                case PlayerFinishesFastAttack:
                case PlayerFinishesInitialFastAttack:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender finishes fast attack");
                    break;
                case PlayerFinishesSpecialAttack:
                case PlayerFinishesInitialSpecialAttack:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender finishes special attack");
                    break;
                case PlayerFinishesTransform:
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender finishes transform");
                    break;
                }

                /* defender performs next action */
                switch (playerEvent) {
                case PlayerStartsTransform:
                    /* defender starts transform */
                    defenderEnergy = Min(defenderEnergy + transformEnergy, maxDefenderEnergy);
                    defenderEventQueue.Add(transformDamageStart, PlayerLandsTransform);
                    defenderEventQueue.Add(transformDuration, PlayerFinishesTransform);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender starts transform");
                    break;
                case PlayerStartsAttack:
                case PlayerStartsInitialAttack:
                    /* defender starts next attack */
                    if (defenderEnergy >= -defenderSpecialAttackEnergy) {
                        /* defender often defers special attacks */
                        if (randomness) {
                            /* random behavior */
                            if ((double) rand() / (RAND_MAX + 1) > defensiveSpecialAttackProbability) {
                                specialAttack = true;
                            } else {
                                specialAttack = false;
                                LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy,
                                         "defender defers special attack");
                            }
                        } else {
                            /* expected behavior */
                            numDefensiveSpecialAttackOpportunities = numDefensiveSpecialAttackOpportunities + 1;
                            if (numDefensiveSpecialAttackOpportunities > numDefensiveSpecialAttackDeferrals) {
                                numDefensiveSpecialAttackOpportunities = 0;
                                specialAttack = true;
                            } else {
                                specialAttack = false;
                                LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy,
                                         "defender defers special attack");
                            }
                        }
                    } else {
                        specialAttack = false;
                    }
                    if (specialAttack) {
                        /* special attack */
                        defenderEnergy = defenderEnergy + defenderSpecialAttackEnergy;
                        defenderEventQueue.Add(defenderSpecialAttackDamageStart, PlayerLandsSpecialAttack);
                        if (playerEvent == PlayerStartsInitialAttack) {
                            defenderEventQueue.Add(defenderSpecialAttackDuration, PlayerFinishesInitialSpecialAttack);
                        } else {
                            defenderEventQueue.Add(defenderSpecialAttackDuration, PlayerFinishesSpecialAttack);
                        }
                        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender starts special attack");
                    } else {
                        /* fast attack */
                        defenderEnergy = Min(defenderEnergy + defenderFastAttackEnergy, maxDefenderEnergy);
                        defenderEventQueue.Add(defenderFastAttackDamageStart, PlayerLandsFastAttack);
                        if (playerEvent == PlayerStartsInitialAttack) {
                            defenderEventQueue.Add(defenderFastAttackDuration, PlayerFinishesInitialFastAttack);
                        } else {
                            defenderEventQueue.Add(defenderFastAttackDuration, PlayerFinishesFastAttack);
                        }
                        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender starts fast attack");
                    }
                    break;
                case PlayerFinishesFastAttack:
                case PlayerFinishesSpecialAttack:
                    /* defender does nothing for a while and then starts next attack */
                    if (randomness) {
                        interval = RandomInterval(defensiveInterval, defensiveIntervalRandomness);
                    } else {
                        interval = defensiveInterval;
                    }
                    defenderEventQueue.Add(interval, PlayerStartsAttack);
                    LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "defender idles");
                    break;
                case PlayerFinishesInitialFastAttack:
                case PlayerFinishesInitialSpecialAttack:
                case PlayerFinishesTransform:
                    /* initial attacks do not start new attacks */
                    break;
                }
            }
        }
        LOGEVENT(logFile, battleTimer, attackerBattleHP, attackerEnergy, defenderBattleHP, defenderEnergy, "battle ends");
        LOGNEWLINE(logFile);

        if (defenderBattleHP <= 0) {
            /* attacker won */
            ++numWins;
        } else {
            /* defender won */
        }
    }

    FREE(15, &levelsRange, &speciesRange, &moveSetsRange,
             &attackerFastAttackName, &attackerFastAttackType, &attackerSpecialAttackName, &attackerSpecialAttackType,
             &defenderFastAttackName, &defenderFastAttackType, &defenderSpecialAttackName, &defenderSpecialAttackType,
             &attackerType1, &attackerType2, &defenderType1, &defenderType2);
    delete[] defensiveInitialIntervals;
    CLOSELOG(logFile);

    /* return probability of attacker winning */
    return (double) numWins / numTrials;
}


double WINAPI DefenderSpeciesAverage(long attackerMoveSetNum, long defenderMoveSetNum)
{
#pragma EXPORT
    XLOPER12   moveSetMatchupsRange;
    XLOPER12   moveSetMatchupAttackersRange;
    XLOPER12   moveSetMatchupDefendersArray;
    int        rowNum, colNum;
    double     matchupSum;
    int        matchupCount;
    int        defenderPokedexNum;
    LPXLOPER12 defender;

    /* do not execute from dialog box */
    if (CalledFromExcelDialog()) return 0.0;

    moveSetMatchupsRange = GetNamedRange(L"\043'Move Set Matchups'!MoveSetMatchups");
    moveSetMatchupAttackersRange = GetNamedRange(L"\053'Move Set Matchups'!MoveSetMatchupAttackers");
    moveSetMatchupDefendersArray = GetNamedArray(L"\053'Move Set Matchups'!MoveSetMatchupDefenders");
    assert(moveSetMatchupDefendersArray.val.array.rows == 1);
    rowNum = Match(attackerMoveSetNum, moveSetMatchupAttackersRange, 0);
    colNum = 1;
    matchupSum = 0;
    matchupCount = 0;
    defenderPokedexNum = defenderMoveSetNum / 1000000;
    for (colNum = 1, defender = moveSetMatchupDefendersArray.val.array.lparray; colNum <= moveSetMatchupDefendersArray.val.array.columns; ++colNum, ++defender) {
        assert(defender->xltype == xltypeNum);
        if ((int) defender->val.num / 1000000 == defenderPokedexNum) {
            matchupSum += IndexNumber(moveSetMatchupsRange, rowNum, colNum);
            ++matchupCount;
        }
    }
    FREE(3, &moveSetMatchupsRange, &moveSetMatchupAttackersRange, &moveSetMatchupDefendersArray);
    return matchupSum / matchupCount;
}
