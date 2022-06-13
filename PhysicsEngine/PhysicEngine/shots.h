#pragma once

#include <wtypes.h>
#include "vector.h"

struct Shot
{
	Vector position;
	Vector oldPosition;
	Vector velocity;
	bool exists;
	bool fromVeg;
	short age;
	short parent;
	short type;
	float value;
	BSTR sourceSpecies;
};

extern std::unique_ptr<Shot[]> shotArray;