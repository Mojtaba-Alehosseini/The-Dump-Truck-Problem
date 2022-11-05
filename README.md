# The-Dump-Truck-Problem
Computer Simulation. The dump truck problem
+ 
+ Example of Discrete Event System Simulation (Fifth Edition) book.
+
+ Six dump trucks are used to haul coal from the entrance of a small mine to the railroad. Each truck is loaded by one of two loaders. After a loading, the truck immediately moves to the scale, to be weighed as soon as possible. Both the loaders and the scale have a first-come-first-served waiting line (or queue) for trucks. Travel time from a loader to the scale is considered negligible. After being weighed, a truck begins a travel time (during which time the truck unloads) and then afterward returns to the loader queue.

+ The distributions of loading time, weighing time, and travel time are given in Tables 3.3, 3.4 and 3.5, respectively. These activity times are generated in exactly the same manner as service times in Section 2.1.5, Example 2.2, using the cumulative probabilities to divide the unit interval into subintervals whose lengths correspond to the probabilities of each individual value. As before, random numbers come from Table A.1 or one of Excel’s random number generators. A random number is drawn and the interval it falls into determines the next random activity time.

+ The purpose of the simulation is to estimate the loader and scale utilizations (percentage of time busy). The model has the following components:

+ System state

+ [LQ(t), L(t), WQ(t), W(t)e], where

+ LQ(t) = number of trucks in loader queue;

+ L(t) = number of trucks (0, 1, or 2) being loaded

+ WQ(t) = number of trucks in weigh queue;

+ W(t) = number of trucks (0 or 1) being weighed, all at simulation time

+ Entities

+ The six dump trucks (DT1,...,DT6).

+ Event notices

+ (ALQ, t , DTi), DTi arrives at loader queue (ALq) at time t;

+ (EL, t , DTi), DTi ends loading (EL) at time t;

+ (EW, t , DTi), DT ends weighing (EW) at time t.

