#pragma once


/* enable or disable multithreaded calculation */
/* VBA callbacks are not thread safe */
#define THREADSAFE 0


/* include or exclude log code */
#define LOG 0


const int dittoPokedexNum = 132;


/* tolerance for floating point inaccuracy */
const double tolerance = 0.0001;


/* player events */
enum PlayerEvents {
    NullEvent,
    PlayerStartsAttack,
    PlayerStartsInitialAttack,
    PlayerStartsTransform,
    PlayerFinishesLongPress,
    PlayerLandsFastAttack,
    PlayerLandsSpecialAttack,
    PlayerLandsTransform,
    PlayerFinishesFastAttack,
    PlayerFinishesInitialFastAttack,
    PlayerFinishesSpecialAttack,
    PlayerFinishesInitialSpecialAttack,
    PlayerFinishesTransform
};


struct EventRecord {
    int          time;
    PlayerEvents event;
};
