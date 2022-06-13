#include "pch.h"
#include "bot.h"

void __stdcall UpdateBotPosition(Bot* bot, float maxSpeed, BOOL zeroMomentum)
{
	if (bot->Fixed == FALSE)
	{
		bot->Velocity += bot->IndependentImpulse / (max(bot->Mass + bot->AddedMass, MinimumMass));

		float actualSpeedSquared = bot->Velocity.MagnitudeSquared();
	
		float maxSpeedSquare = powf(maxSpeed, 2);

		if (actualSpeedSquared > maxSpeedSquare)
			bot->Velocity = bot->Velocity.Unit() * maxSpeed;

		bot->Position += bot->Velocity;

		if (zeroMomentum == TRUE)
			bot->Velocity = { 0,0 };
	}
	else
	{
		bot->Velocity = { 0, 0 };
	}
}