#include "pch.h"
#include <wtypes.h>
#include <memory>
#include "shots.h"

constexpr size_t StartingArraySize = 50;

size_t currentSize = StartingArraySize;
std::unique_ptr<Shot[]> shotArray(new Shot[currentSize]);

void InitShots()
{
	shotArray.reset(new Shot[currentSize]);

	for (size_t i = 0; i < currentSize; i++)
		shotArray[i] = Shot();
}

int NewShot(short parent, short type, float value, float rangeMultiplier, BOOL offset, BSTR species, BOOL fromVeg)
{
	size_t a = FirstSlot();

	if (a >= currentSize)
	{
		a = currentSize;
		size_t newSize = currentSize * 1.1;

		std::unique_ptr<Shot[]> newShotArray(new Shot[newSize]);

		for (size_t i = 0; i < newSize; i++)
		{
			if (i < currentSize)
				newShotArray[i] = shotArray[i];
			else
				newShotArray[i] = Shot();
		}
		currentSize = newSize;
		shotArray.reset(newShotArray.release());
	}

	value = min(value, 32000);

	shotArray[a].exists = true;
	shotArray[a].age = 0;
	shotArray[a].parent = parent;
	shotArray[a].value = roundf(value);
	shotArray[a].sourceSpecies = species;
	shotArray[a].fromVeg = fromVeg == TRUE;

	if (type > 0 || type == -100)
		shotArray[a].type = type;
	else
	{
		shotArray[a].type % 8;
		if (shotArray[a].type == 0) shotArray[a].type = -8;
	}
}

size_t FirstSlot()
{
	size_t i = 0;

	while (shotArray[i].exists && i < currentSize)
		i++;

	return i;
}