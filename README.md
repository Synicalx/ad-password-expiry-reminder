# ad-password-expiry-reminder
Rough script to email AD users a day before their password expires.

Treat this as a template, that's probably not very well made. Please test this on a test environment prior to using it in production!


# Overview

This script takes an OU, and recursively grabs individual User objects and their "msDS-UserPasswordExpiryTimeComputed" attribute. From this attribute, it converts it to standard DateTime and compares that to a somewhat poorly calculate 'next business day' with both variables being compared on their 'DayofYear' value. If they match, the user is added to a List and then emailed a notification using a responsive HTML email template. 

There is logging for each step, that will either log success or failure with a timestamp.

# Calculating NBD

Very simple method of doing this is used in the script, if it's a weekday other than Friday then the NBD is one day away, but if it's a Friday then the NBD is 3 days away (ie Monday). Naturally this doesn't handle public holidays, which you may want to add for your country/state/territory's particular public holidays.

# Determining WHEN to notify users

There are a number of ways to do this, in the case of this script it simply compares the password expiry to a calculated date (NBD). You may also wish to compare the age of a user's password to the domain's password policy, which may be slightly more efficient in some cases.
