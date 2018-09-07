#include <assert.h>
#include <limits.h>

#include "EventQueue.h"


EventQueue::EventQueue(void)
{
    queue = nullptr;
    numEntries = 0;
    maxEntries = 0;
}


EventQueue::~EventQueue(void)
{
    delete[] queue;
}


void EventQueue::Initialize(int numInitialIntervals)
{
    if (maxEntries != numInitialIntervals * 2 + 1) {
        maxEntries = numInitialIntervals * 2 + 1;
        if (queue) {
            delete[] queue;
        }
        queue = new EventRecord[maxEntries];
    }
    queue[0].time = INT_MAX;
    queue[0].event = NullEvent;
    numEntries = 0;
}


int EventQueue::Timer(void)
{
    assert(queue[0].time < INT_MAX);
    return queue[0].time;
}


void EventQueue::Add(int time, PlayerEvents playerEvent)
{
    int i;

    assert(time < INT_MAX);
    assert(numEntries >= 0);
    assert(numEntries < maxEntries);
    i = numEntries;
    /* find position for new entry and shift queue backward to make room */
    while (i >= 0 && queue[i].time > time) {
        queue[i + 1] = queue[i];
        --i;
    }
    /* insert new entry */
    ++i;
    queue[i].time = time;
    queue[i].event = playerEvent;
    ++numEntries;
}


void EventQueue::CountDown(int time)
{
    int i;

    for (i = 0; i < numEntries; ++i) {
        queue[i].time -= time;
    }
}


PlayerEvents EventQueue::Pop(void)
{
    PlayerEvents event;
    int          i;

    assert(queue[0].time == 0);
    /* get first event */
    event = queue[0].event;
    /* shift queue forward */
    for (i = 0; i < numEntries; ++i) {
        queue[i] = queue[i + 1];
    }
    --numEntries;
    assert(numEntries >= 0);
    return event;
}
