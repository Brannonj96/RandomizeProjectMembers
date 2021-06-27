// randomizeGroups reads from the current spreadsheet (wherever the button is) and checks the 
// ProjectList sheet for correctly formatted values and if present, reads them and randomly
// assigns members into the project based on their preferences

class Member {
  constructor(name, preferences) {
    this.name = name
    this.preferences = preferences
  }
}


function randomizeGroups() {
  // Load the ui for use in giving error alerts
  var ui = SpreadsheetApp.getUi()

  // Load the spreadsheet that the button/caller was in
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  // Check the max group size cell for any errors
  var maxGroupSize = spreadsheet.getSheets()[0].getRange(8,2).getValue()
  if (!Number.isInteger(maxGroupSize) || maxGroupSize<=0) {
    ui.alert("Max group size must be set to a whole number greater than 1. Please set it in cell B8")
    return
  }

  var minGroupSize = spreadsheet.getSheets()[0].getRange(9,2).getValue()
  if (!Number.isInteger(minGroupSize) || minGroupSize<0 || minGroupSize>maxGroupSize) {
    ui.alert("Min group size must be set to a whole number ranging from 0 to the Maximum Group Size from cell B8. Please set it in cell B9")
    return
  }

  // Load the sheet ProjectList that is expected to be formatted as shown in the Main sheet
  // Errors if it does not exist
  var dataSheet = spreadsheet.getSheetByName("ProjectList")
  if (!dataSheet) {
    ui.alert("The sheet \"ProjectList\" must exist to run. Please create it and try again")
    return
  }
  
  // Load the data off the project sheet. If there are no members present, it gives an error. If there are no projects present, it gives an error.
  var data = dataSheet.getDataRange().getValues()
  if (data.length<=1 || data[0].length<=1) {
    ui.alert("Please make sure the ProjectList sheet is formatted correctly as seen in the main sheet and that at least one member and one project is present with preferences")
    return
  }

  // Load the first row, which should be project names. Slice to after the first element, which should be the "Name" column tab instead of a project name
  var rawProjects = data[0].slice(1)

  // Filter out any empty strings from project names. These are present if there values in columns past the last project's column
  var projects = rawProjects.filter(function (value) {
    return value != ""
  })

  if (minGroupSize * projects.length > data.length-1) {
    ui.alert("The minimum group size is impossible to achieve. Please decrease it in cell B9 or add more members to the ProjectList sheet")
    return
  }

  // Creates a map in the form {ProjectName: MemberArray} e.g. {Project1: [member1, member2, member3], Project2: [member4,member5]}
  // Returns an error on duplicate project names.
  // This is used to assign members to projects later. Every project is initialized with empty membership arrays
  var projectsAndGroups = new Map()
  for (var i = 0; i<projects.length; i++) {
    if (projectsAndGroups.has(projects[i])) {
      ui.alert(`${projects[i]} is a duplicate. Please make sure every project name is unique.`)
      return
    }
    projectsAndGroups.set(projects[i], [])
  }

  // Begin iterating over the rows of students and their preferences
  // Create a list of the Member class. Member will have a name and a preference list
  //
  // Note the preference list is represented in ascending order so that they may be popped by preference priority in the correct order.

  var members = []
  for (var i = 1; i<data.length; i++) {
     var row = data[i]
     if (row[0] == "") {
       ui.alert(`The name cannot be blank. Correct row ${i+1}`)
       return
     }

     var memberName = row[0]
     var orderedPreferences = new Array(projects.length)
     // Get all the preference values which should form a complete range 1 .. N though not necessarily in order.
     // If the length of preferences is not the same as the project length it errors. Removed empty slots from the project array to check this here.
     var preferences = row.slice(1)
     if (preferences.length != projects.length) {
       ui.alert(`${row[0]} does not have the proper number of preferences. Please make sure each project has a preference and there are no extra values`)
     }

     // Initialize a set to keep track of duplicate preferences  
     var seenPreferences = new Set()

     for (var j = 0; j < preferences.length; j++) {
       var memberProjectPreference = preferences[j]

       // Error checking on preferences. If it's not in the range 1..N or isn't an integer, it errors and says the row it errors on.
       if (!Number.isInteger(memberProjectPreference) || memberProjectPreference<1 || memberProjectPreference>projects.length) {
         ui.alert(`${row[0]} has inserted an invalid preference. Please make sure all preferences are whole numbers in the range 1 .. N`)
         return
       }

       // Error checking for duplicates and says row it errored on
       if (seenPreferences.has(memberProjectPreference)) {
         ui.alert(`${row[0]} has duplicate preference values. Please make sure all preferences are unique whole numbers in the range 1 .. N`)
         return
       }
       seenPreferences.add(memberProjectPreference)

       // Read the preference of the j'th project (left to right) and then assign that project to the index of the preference in reverse order.
       // E.g. If we had project1, project2, project 3 and preferences 2 3 1, then the preference array would end up looking like
       // [project2, project1, project3] because projects and preferences are read left to right, and then the preference itself
       // is used to determine the index to place them (in ascending order). Project2 is first as least desirable, then project3, then project 1 as most desirable.
       var projectPreference = preferences.length - memberProjectPreference
       orderedPreferences[projectPreference] = projects[j]
     }
     members.push(new Member(memberName, orderedPreferences))
  }

  // A random member is chosen each iteration and then slotted based on their preference.
  // if their first preference has available slots, they are placed there, if not then it moves to the 
  // 2nd preference and so on until the Nth preference. If they are not slotted because all groups are full, it errors.
  //
  // On each iteration, a new random member is chosen.
  for (var i = 0; i<members.length; i++) {
    
    // Gives a random index in range 0 .. M where M is the remaining members to be slotted
    var randomIndex = Math.floor(Math.random() * (members.length - i))

    // This is an in-place random choice method; choose a random index 0..M where M is the remaining items to check
    // once you get that index, store the value at the index, replace the index with the Mth value, and replace the Mth value
    // with the stored value from the index
    //
    // Example:
    // [Member1, Member2, Member3, Member4, Member5] Random Number [0,5) chooses 3 and selects Member4 for placement, then the array becomes
    // [Member1, Member2, Member3, Member5, Member4] Random Number [0,4) chooses 1 and selects Member2 then the array becomes
    // [Member1, Member5, Member3, Member2, Member4] Random Number [0,3) chooses 0 and selects Member1 then the array becomes
    // [Member3, Member5, Member1, Member2, Member4] Random Number [0,2) chooses 1 and selects Member5 then the array becomes
    // [Member3, Member5, Member1, Member2, Member4] note it has not changed because the last step choose the last index, 1, in the range 0..M. Random Number [0,1) chooses 0 and selects Member3
    // [Member3, Member5, Member1, Member2, Member4] note it did not change because it selected the last index in the range 0..M. The resulting array is also a reverse of the order they were chose
    var member = members[randomIndex]
    members[randomIndex] = members[members.length - 1 - i]
    members[members.length - 1 - i] = member

    // Iterate over the member's preferences and fit them into the first available project. They are placed into the previously created
    // projectsAndGroups map.
    // If all projects are full, it returns an error.
    while (member.preferences.length > 0) {
      var slotted = false
      var preference = member.preferences.pop()

      if (projectsAndGroups.get(preference).length < maxGroupSize) {
          projectsAndGroups.get(preference).push(member)
          slotted = true
          break
      }
    }
    
    if (!slotted) {
      ui.alert(`Unable to fit all members into projects. Please adjust the maximum group size or the number of members`)
      return
    } 
  }

  // Check if any group is below minimum size
  var belowMinimumSizeProjects = new Set()
  for (let [project, group] of projectsAndGroups) {
    if (group.length < minGroupSize) {
      belowMinimumSizeProjects.add(project)
    }
  }

  // If a group is below the minimum size, choose a random project that has more than the minimum
  // then choose random members out of that project. If their next highest preference
  // is a project below the minimum threshold, then remove them from their current project
  // and add the member to their next preference, the one below the minimum threshold
  // if the threshold is then met, remove it from the set containing the groups below minimum size
  //
  // If no member in the random project has a next preference that is below the threshold, then
  // it selects another project (without repeating) until a member with that preference is found
  //
  // If after every project has been iterated, it loops through the project again to check the next preference
  // of every member until all groups have at least the minimum size
  while (belowMinimumSizeProjects.size > 0) {
    for (var i = 0; i<projects.length; i++) {

      // See the above loop that initally places members for an explanation on this random-choice method
      var randomIndex = Math.floor(Math.random() * (projects.length - i))
      var nextProject = projects[randomIndex]
      projects[randomIndex] = projects[projects.length - 1 - i]
      projects[projects.length - 1 - i] = nextProject

      if (belowMinimumSizeProjects.has(nextProject)) {
        continue
      }
      var groupMembers = projectsAndGroups.get(nextProject)
      for (var j = 0; j < groupMembers.length; j++) {

        // See the above loop that initally places members for an explanation on this random-choice method
        var randomMemberIndex = Math.floor(Math.random() * (groupMembers.length - j))
        var nextMember = groupMembers[randomMemberIndex]
        groupMembers[randomMemberIndex] = groupMembers[groupMembers.length -1 - j]
        groupMembers[groupMembers.length - 1 - j] = nextMember

        var memberNextPreference = nextMember.preferences.pop()
        if (belowMinimumSizeProjects.has(memberNextPreference)) {
          projectsAndGroups.get(memberNextPreference).push(nextMember)
          groupMembers.splice(groupMembers.length - 1 - j, 1)

          if (projectsAndGroups.get(memberNextPreference).length == minGroupSize) {
            belowMinimumSizeProjects.delete(memberNextPreference)
          }
        }
      }
    }
  }
  
  // Attempt to get the spreadsheet called results
  var results = spreadsheet.getSheetByName("Results")

  //If it is null, create it
  if (!results) {
    results = spreadsheet.insertSheet()
    results.setName("Results")
  }
  // Set it as the active sheet for writing
  spreadsheet.setActiveSheet(results)
  // Set it to the last sheet position for consistency
  spreadsheet.moveActiveSheet(spreadsheet.getNumSheets())
  // Clear it of any previous values
  results.clear()

  // Write the projects with their group members where each project is a column header and group members are under that e.g.
  // Project1 Project2  Project3
  // Member1  Member2   Member3
  // Member4  Member5   Member6
  var col = 1
  var defaultColumnSize = results.getColumnWidth(col)
  for (let [project, members] of projectsAndGroups) {
    var row = 1
    results.getRange(row++, col).setValue(project)
    for (var i = 0; i<members.length; i++) {
      results.getRange(row++, col).setValue(members[i].name)
    }

    // Autoresize and add a completely arbitrary number of pixels that I think looks alright
    // If it's below the default size, just set it to the default size.
    results.autoResizeColumn(col)
    results.setColumnWidth(col, results.getColumnWidth(col)+25)
    if (results.getColumnWidth(col) < defaultColumnSize) {
      results.setColumnWidth(col, defaultColumnSize)
    }
    col++
  }
}

