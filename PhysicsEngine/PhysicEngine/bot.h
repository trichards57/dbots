#pragma once

#include "vector.h"

constexpr float MinimumMass = 0.25;

struct Bot {
	Vector Position;
	Vector Velocity;
	Vector IndependentImpulse;
	BOOL Fixed;
	float Mass;
	float AddedMass;
};

extern "C" {
	extern _declspec(dllexport) void __stdcall UpdateBotPosition(Bot* bot, float maxSpeed);
}