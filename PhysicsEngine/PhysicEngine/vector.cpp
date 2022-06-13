#include "pch.h"
#include "vector.h"
#include "shared.h"

Vector Vector::operator/(const float scalar) const
{
	return { x / scalar, y / scalar };
}

Vector Vector::operator*(const float scalar) const
{
	return { x * scalar, y * scalar };
}

Vector Vector::operator+(const Vector v2) const
{
	return { x + v2.x, y + v2.y };
}

void Vector::operator+=(const Vector v2)
{
	x += v2.x;
	y += v2.y;
}

float Vector::MagnitudeSquared() const
{
	float actualX = x;
	float actualY = y;
	if (fabsf(x) > MaxInt)
		actualX = copysignf(MaxInt, x);
	if (fabsf(y) > MaxInt)
		actualY = copysignf(MaxInt, y);

	return actualX * actualX + actualY * actualY;
}

float Vector::Magnitude() const
{
	float minVal = min(fabs(x), fabs(y));
	float maxVal = max(fabs(x), fabs(y));

	if (maxVal < 0.00001)
		return 0;
	return maxVal * sqrtf(1 + powf(minVal / maxVal, 2));
}

float Vector::InverseMagnitude() const
{
	float magnitude = Magnitude();

	if (magnitude == 0)
		return -1;
	return 1 / magnitude;
}

Vector Vector::Unit() const
{
	return (*this) * InverseMagnitude();
}