#pragma once


#include "BattleSimulator.h"


class EventQueue {
public:
                 EventQueue  (void);

                 ~EventQueue (void);

    void         Initialize  (int numInitialIntervals);

    int          Timer       (void);

    void         Add         (int time, PlayerEvents playerEvent);

    void         CountDown   (int time);

    PlayerEvents Pop         (void);

private:
    EventRecord *queue;
    int         numEntries;
    int         maxEntries;
};
