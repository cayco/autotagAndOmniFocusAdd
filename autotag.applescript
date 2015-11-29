(*
	SmartTagMail script
	Paul Hagstrom, January 2008
	Krzysztof Kajkowski, November 2015
	
	Usage:
	Set autoTagTrigger and autoBoxTrigger to something, e.g., "autotag:" and "autobox:"
	
	Requires (of course): MailTags.
	
	In the notes of an address book entry, you can follow the trigger with a word.
	When this rule is invoked on a message, it will look in associated address book records,
	and, where it finds the trigger in the notes, it will tag/move the message accordingly.
	A tag will be added for each person with a trigger associated with the message.
	If one or more of the people is a member of a group whose name contains the
	autotag trigger, that tag will be added as well.  (e.g. "UIS autotag:uis")
	Where there are multiple autobox triggers encountered, the first one found
	will be acted on UNLESS a later one starts with "!", in which case it will
	be then considered to be the first one found.  This, the message is moved to
	the first box found, or, if there are any forced boxes, then the last forced
	box found.
	
	autotag takes a tag name, autobox takes a mailbox name.
	Mailbox names need to reflect the hierarchy, e.g. Lists/scriptlists
	It will look in the note field from trigger: until the end of its line
	I envision using this myself as a script manually called using Mail Act-On,
	but if you really come to trust it, I suppose you could automatically apply it
	to all incoming mail.
	
	autoproject is not implemented yet, but should be easy.	
	
	A note about case.  Email addresses in the address book need to be all
	lowercase to be found.  You have control over the Address Book.  Fix broken ones.
	You don't have control over incoming email, so those are lowercased for you.
	It's a flaw in Address Book that there's no way to search for email addresses
	case insensitively.
	
	If this is called from a rule with "sender only" in the name, then only
	the sender, and not the recipients, will be scanned.  (Intended use is
	for mail sent to big lists from a known correspondent.) In case it is useful, it
	will also look for "recipients only" in the name, which will likewise
	trigger scanning of the recipients and not the sender. And, finally, it will
	look for "addressees only" in the name, which will scan the to:
	recipients, but not the cc: recipients. The priority is sender, recipients, addressees.
	If a rule name contains more than one of these keywords, only the first priority one takes effect.
*)

property autoTagTrigger : "autotag:" --text for tag trigger in notes and group names
property autoBoxTrigger : "autobox:" --text for box trigger in notes and group names
property autoProjTrigger : "autoproject:" --text for project trigger in notes and group names

property debugLevel : 2 --set from 0 to 3 depending on how detailed you want console output to be
property dryRun : false --set to false to actually move and tag, true to just pretend
property triggerNames : {"tag", "box"} --used in logging for findTrigger at debugLevel 3

(* Send debugging information to the console *)
on Logger(level, str)
	if debugLevel > level then
		do shell script "logger " & quoted form of ("SmartTagMail - " & str)
	end if
end Logger

(* Collect emails from the current message and return them in a list *)
(* the parameters govern whether the sender's email is collected,
whether all recipients' emails are collected, and whether just the to: recipients are collected.
Note that if scanRecipients is false, the value of scanTo is irrelevant. *)
on collectEmails(msg, scanSender, scanRecipients, scanTo)
	my Logger(2, "Collecting emails.")
	using terms from application "Mail"
		set theEmails to {}
		--first check the sender, if we are supposed to
		if scanSender then
			set theSender to sender of msg
			set theEmail to extract address from theSender
			set theEmails to {theEmail}
			my Logger(1, "Collect emails: Sender: " & theEmail)
		end if
		--now go through the recipients, if we are supposed to
		if scanRecipients then
			if scanTo then
				set theRecipients to every to recipient of msg
			else
				set theRecipients to every recipient of msg
			end if
			repeat with theRecipient in theRecipients
				set theEmail to address of theRecipient
				set theEmails to theEmails & {theEmail}
				my Logger(1, "Collect emails: Recipient: " & theEmail)
			end repeat
		end if
	end using terms from
	return theEmails
end collectEmails

(* Scan the note for triggers passed in as the second parameter.  Multiple hits are possible, but each trigger takes the rest of its line. That is, you can have autotag:X and autotag:Y on two different lines and add both tags X and Y. The way it parses it that it cuts out everything up to the trigger and the processes the rest again. What this means is that you'll get funny results if your trigger is "autotag:" and you try to use it to tag a message with the tag "autotag:"  Don't do that. *)
on findTriggers(theNote, theTriggers)
	my Logger(2, "Scanning for triggers.")
	set foundTriggers to {}
	repeat with i from 1 to length of theTriggers
		set theTrigger to item i of theTriggers
		set theTail to theNote
		set theResults to {}
		repeat
			set theOffset to offset of theTrigger in theTail
			if theOffset > 0 then
				set theOffset to theOffset + (length of theTrigger)
				set theTail to (text theOffset thru (length of theTail) of theTail) as text
				set theValue to paragraph 1 of theTail
				set theResults to theResults & theValue
				my Logger(2, "Trigger for  " & (item i of triggerNames) & ": " & theValue)
			else
				exit repeat
			end if
		end repeat
		set foundTriggers to foundTriggers & {theResults}
	end repeat
	return foundTriggers
end findTriggers

on processPerson(theEmail, triggerList)
	tell application "Contacts"
		
		--look for a person who has this email address (see note at top about case)
		try
			set foundPerson to (first person where value of every email of it contains theEmail)
		on error
			my Logger(1, "Scan Address Book: No entry for " & theEmail)
			return {}
		end try
		try
			--scan the person's note for triggers
			set theNote to (get note of foundPerson)
			set theName to (get name of foundPerson)
			my Logger(1, "Scan Address Book: Processing " & theName & " - " & theEmail)
			set foundTriggers to (my findTriggers(theNote, triggerList))
			--Look for groups that might contain additional triggers
			set theGroups to every group of foundPerson
			repeat with theGroup in theGroups
				set statusString to ""
				set theGroupName to name of theGroup
				my Logger(2, "Scan Address Book: Processing group " & theGroupName)
				set groupTriggers to (my findTriggers(theGroupName, triggerList))
				repeat with i from 1 to length of groupTriggers
					set foundItems to (a reference to item i of foundTriggers)
					set contents of foundItems to contents of foundItems & item i of groupTriggers
				end repeat
			end repeat
			return foundTriggers
		on error errMsg number errNumber
			my Logger(1, "Scan Address Book: Error for: " & theEmail & ": " & errMsg)
			return {}
		end try
	end tell
end processPerson

using terms from application "Mail"
	on perform mail action with messages theMessages for rule theRule
		--here is where the mapping from triggers to tag/box/project happens. Most of the rest of the code is pretty general.
		set triggerList to {autoTagTrigger, autoBoxTrigger}
		--set triggerList to {autoTagTrigger, autoBoxTrigger, autoProjTrigger}
		set tagItem to 1
		set boxItem to 2
		--set projectItem to 3
		Logger(0, "***Starting: triggers " & triggerList)
		Logger(2, "***Starting: theRule name is " & name of theRule)
		set scanSender to true
		set scanRecipients to true
		set scanTo to false
		if name of theRule contains "sender only" then
			set scanRecipients to false
			Logger(1, "***Scanning only sender (rule name contains sender)")
		else
			if name of theRule contains "recipients only" then
				Logger(1, "***Scanning only recipients")
				set scanSender to false
			else
				if name of theRule contains "addressees only" then
					Logger(1, "***Scanning only to: recipients")
					set scanTo to true
				else
					Logger(1, "***Scanning sender and all recipients")
				end if
			end if
		end if
		if dryRun then
			Logger(0, "***DRY RUN")
		end if
		Logger(2, "***DEBUG LEVEL: " & debugLevel)
		Logger(2, "Messages selected: " & (length of theMessages))
		repeat with msg in theMessages
			Logger(2, "Beginning message processing.")
			set theEmails to my collectEmails(msg, scanSender, scanRecipients, scanTo)
			set combinedTriggers to {}
			repeat with theEmail in theEmails
				set foundTriggers to processPerson(theEmail, triggerList)
				if length of foundTriggers > 0 then --if it isn't then the person wasn't in the address book, ignore
					if length of combinedTriggers is 0 then --this is the first substantive time through the loop
						set combinedTriggers to foundTriggers
					else
						repeat with i from 1 to length of foundTriggers
							set combinedItems to (a reference to item i of combinedTriggers)
							set contents of combinedItems to contents of combinedItems & item i of foundTriggers
						end repeat
					end if
				end if
			end repeat
			(* Having found all of the available triggers, we now deal with them.  Tags first. *)
			Logger(2, "Consolidating tags.")
			using terms from application "MailTagsHelper"
				set newTags to keywords of msg
				Logger(2, "Existing tags: " & (newTags as text))
				set tagsDirty to false
				repeat with theTag in item tagItem of combinedTriggers
					if length of theTag > 0 then
						if newTags does not contain theTag then
							set newTags to newTags & theTag
							-- Exclude @Waiting tag to not add this context in OmniFocus
							if theTag does not contain "@Waiting" then
								-- this is special tag for OmniFocus task creation
								set OFTag to theTag
							end if
							set tagsDirty to true
							
						end if
					end if
				end repeat
			end using terms from
			(* Now, deal with the boxes. *)
			Logger(2, "Consolidating boxes.")
			set moveBox to ""
			repeat with theBox in item boxItem of combinedTriggers
				if character 1 of theBox is "!" then
					set testBox to (rich text 2 thru (length of theBox) of theBox) as rich text
					set force to true
				else
					set testBox to theBox
					set force to false
				end if
				if exists mailbox testBox then
					if force then
						Logger(2, "Mailbox forced to: " & testBox)
						set moveBox to testBox
					else
						if length of moveBox is 0 then
							Logger(2, "Mailbox set to: " & testBox)
							set moveBox to testBox
						else
							Logger(2, "Mailbox ignored: " & testBox)
						end if
					end if
				else
					--if the mailbox doesn't exist, ignore it
					Logger(2, "Mailbox does not exist: " & testBox)
					set testBox to ""
				end if
			end repeat
			(* Later I will add project handling here too. It will work just like Boxes, there's only one. *)
			(* Now, process the message. Because doing multiple things to a message can cause it to get lost, I will use the workaround proposed by ahmontgo on the indev.ca forum, and move the message first, then find it again, and perform the other operations *)
			set msgID to the message id of msg
			--Move if needed
			if length of moveBox > 0 then
				if mailbox of msg is mailbox moveBox then
					Logger(0, "Already in target box: " & moveBox)
				else
					Logger(0, "Moving to box: " & moveBox)
					if not dryRun then
						set mailbox of msg to mailbox moveBox
						--if we need to do more, then find the message again post-move
						if tagsDirty then
							set targetMessages to (messages of mailbox moveBox whose message id is msgID)
							set msg to the first item of targetMessages
						end if
					end if
				end if
			end if
			--Tag if needed
			using terms from application "MailTagsHelper"
				if tagsDirty then
					Logger(0, "Setting keywords to: " & (newTags as text))
					if not dryRun then
						set keywords of msg to newTags
					end if
				end if
			end using terms from
			-- create a task in OmniFocus
			
			
			tell application "OmniFocus"
				log "OmniFocus calling process_message in MailAction script"
			end tell
			set theSubject to subject of msg
			set singleTask to false
			if (theSubject starts with "Fwd: ") then
				-- Whole forwarded messages shouldn't split.
				set singleTask to true
				set theSubject to rich text 6 through -1 of theSubject
			end if
			set theText to "--" & theSubject & (OFTag as rich text) & return & content of msg
			tell application "OmniFocus"
				parse tasks into default document with transport text theText as single task singleTask
			end tell
			
			
			
		end repeat
	end perform mail action with messages
end using terms from