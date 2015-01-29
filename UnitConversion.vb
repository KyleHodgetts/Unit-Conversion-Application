Module UnitConversion
    'Function: A conversion program allowing the user to choose units to convert from three different catagories
    'Astronomical: Light Minutes - Kilometers
    'Force: Kilogram Force - Dyne
    'Pressure: Bar - Dekapascal
    'Author: Kyle Hodgetts
    'Version: 1.0
    Dim validChoice As Boolean      'Variable that will be used in the error handling process to avoid program crashing

    'Starting menu that all users start at.
    'Different key strokes take users to different converters
    Sub Main()
        Dim userChoice As String        'Stores first user input in order to move to the correct subprocedure
        'This sub procedure will loop while invalid input has been entered to avoid program crashing
        'Loop will stop once "X" has been entered
        Do
            Console.Clear()
            Console.WriteLine("   ___                                     _                    ")
            Console.WriteLine("  / __\ ___   _ __ __   __ ___  _ __  ___ (_)  ___   _ __   ___ ")
            Console.WriteLine(" / /   / _ \ | '_ \\ \ / // _ \| '__|/ __|| | / _ \ | '_ \ / __|")
            Console.WriteLine("/ /___| (_) || | | |\ V /|  __/| |   \__ \| || (_) || | | |\__ \")
            Console.WriteLine("\____/ \___/ |_| |_| \_/  \___||_|   |___/|_| \___/ |_| |_||___/")
            Console.WriteLine("                                                                ")
            Console.WriteLine()
            'Indentifies program for user
            Console.WriteLine("This is a conversions program for Force, Pressure and Astronomical Units")
            Console.WriteLine()
            Console.WriteLine("Which kind of units would you like to convert?: ")
            'Prompts user input
            Console.WriteLine("Enter 'A' for Force")
            Console.WriteLine("Enter 'B' for Pressure")
            Console.WriteLine("Enter 'C' for Astronomical")
            Console.WriteLine()
            Console.WriteLine("Alternatively")
            Console.WriteLine("Enter 'H' for the Help Menu")
            Console.WriteLine("Enter 'X' to Exit")
            userChoice = Console.ReadLine()
            userChoice = userChoice.ToUpper

            'Based on input, program will call up one of the following sub procedures
            If userChoice = "A" Then
                forceConversion()

            ElseIf userChoice = "B" Then
                pressureConversion()

            ElseIf userChoice = "C" Then
                astronomicalConversion()

            ElseIf userChoice = "H" Then
                helpMenu()

            ElseIf userChoice = "X" Then
                Console.WriteLine("Program Closing, Press Enter to continue")


            Else
                Console.WriteLine("Invalid input, please enter A, B, C or H or X, Press Enter to try again")
                Console.ReadLine()

            End If

        Loop Until (userChoice = "X")
        Console.ReadLine()

    End Sub

    'Help menu which can only be exited by hitting enter once called up
    'Any other input is irrelavant
    Sub helpMenu()
        Console.Clear()
        'Programs Help menu with the instructions needed to operate the program.
        Console.WriteLine()
        Console.WriteLine("HELP MENU")
        Console.WriteLine()
        Console.WriteLine("To navigate through the program, enter appropriate keys when instructed")
        Console.WriteLine()
        Console.WriteLine("When in a sub menu (not the main menu), press 'R' to return to the main menu.")
        Console.WriteLine()
        Console.WriteLine("I hope this was helpful")
        Console.WriteLine()
        Console.WriteLine("Press the Enter key to return to the main menu")
        Console.ReadLine()

    End Sub

    Sub astronomicalConversion()
        'Light Minutes/Kilometres converter menu
        'From here user can choose which astronomical unit they wish to convert
        Dim userChoice As String        'Stores user input at this menu and calls up appropriate Sub Procedures

        'Loop which will be activated by invalid input
        'Loop will end when input is "R" or "X"
        Do
            Console.Clear()
            Console.WriteLine()
            Console.WriteLine("You have chosen Astronomical Conversion")
            Console.WriteLine("How would you like to convert?")
            Console.WriteLine("Enter 'A' for Light Minutes to Kilometers")
            Console.WriteLine("Enter 'B' for Kilometers to Light Minutes")
            Console.WriteLine()
            Console.WriteLine("Enter R to return to the main menu")
            Console.WriteLine("Enter 'X' to Exit")
            userChoice = Console.ReadLine()
            userChoice = userChoice.ToUpper

            'Based on the input provided by user, different sub procedures will be called up
            'Invalid input will trigger error message and activate loop to ask user to input again
            'Loop will end when "R" or "X" is entered
            If userChoice = "A" Then
                astroLighttoKil()

            ElseIf userChoice = "B" Then
                astroKiltoLight()

            ElseIf userChoice = "R" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            ElseIf userChoice = "X" Then
                End

            Else
                Console.WriteLine("Invalid input, please enter A, B or X. Press Enter to try again")
                Console.ReadLine()
            End If
        Loop Until userChoice = "R" Or userChoice = "X"

    End Sub

    Sub astroLighttoKil()
        'Light Minutes to Kilometres converter
        'Users can enter a Light Minutes figure and have a Kilometres figure returned
        Dim astroInput As Decimal       'Stores light minutes and allows user to enter a decimal figure
        Dim ans As String               'Stores the answer of the user when prompted to convert again
        Dim validAstro As Boolean       'Variable that is used in the error catching process to avoid program crashing

        'Loop that will be exited once input "Y" is entered
        Do
            Console.Clear()

            'Loop that will stop invalid input from crashing the program and prompt user to enter valid input
            Do
                Console.WriteLine("Enter the amount of Light Minutes you wish to convert: ")
                Try
                    validAstro = True
                    astroInput = Console.ReadLine()
                    Console.WriteLine()
                    Console.WriteLine("{0} Light Minutes is equal to {1} Kilometeres", astroInput, convertedLightKil(astroInput))
                Catch
                    validAstro = False
                    Console.WriteLine("Please enter a valid choice")
                End Try
            Loop Until validAstro = True

            Console.WriteLine()
            ' keep asking user if would like.. only y or n allowed
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")

            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to continue")
                Console.ReadLine()
            End If

        Loop While (ans = "Y")

    End Sub

    'The formula that carries out the Light Minutes to Kilometres converion - astro input value passed, returns converted figures
    Function convertedLightKil(ByVal astroInput As Decimal) 'Takes a copy of the user's input without affecting original memory location
        Dim convertedLightToKil As Decimal  'Stores the converted figures allowing for a decimal which will be then be returned to the calling sub procedure

        convertedLightToKil = astroInput * 17987547.48      'Formula for converting Light Minutes to Kilometres

        Return (convertedLightToKil)    'Returns converted figures

    End Function

    'Kilometres to Light Minutes converter
    'Users can enter a Kilometre figure and get it returned in Light Minutes
    Sub astroKiltoLight()
        Dim astro As Decimal        'Stores user input and allows for decimal figures
        Dim ans As String           'Stores response to prompting user to convert again
        Dim validAstro As Boolean   'Variable used in error catching process to avoid program crashing
        Do
            Console.Clear()

            'Loop that will catch input errors and prompt user to input the required input
            Do
                Console.WriteLine("Enter the amount of Kilometeres you wish to convert: ")
                Try
                    validAstro = True
                    astro = Console.ReadLine()
                    Console.WriteLine()
                    Console.WriteLine("{0} Kilometeres is equal to {1} Light Minutes", astro, kilToLight(astro))
                Catch
                    validAstro = False
                    Console.WriteLine("Please enter a valid figure")
                End Try
            Loop Until validAstro = True

            Console.WriteLine()
            'Loop that prompts user to convert again and will loop until input is not "Y"
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")

            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            End If

        Loop While (ans = "Y")
    End Sub

    'The function that carries out the kilometeres to light minutes conversion - astro units value passed, returns
    Function kilToLight(ByVal astro As Decimal)     'Takes a copy of the user's input without affecting original memory location
        Dim convertedKilLight As Decimal            'Stores Converted kilometres as a Decimal                                    '
        convertedKilLight = astro * 0.00000055594015866     'Formula for converting Kilometres into Light Minutes
        Return (convertedKilLight)      'Returns converted figures as LightMinutes
    End Function

    'Force Conversion menu
    'From here users can choose what type of force units they wish to convert
    Sub forceConversion()
        Dim forceInput As String        'Stores user input and calls up the appropriate sub procedure

        'Loop that will loop until input = "R" or "X"
        Do
            Console.Clear()
            Console.WriteLine("You have chosen Force Conversion")
            Console.WriteLine()
            Console.WriteLine("How would you like to convert?")
            Console.WriteLine("Enter 'A' for Kilogram-Force to Dyne")
            Console.WriteLine("Enter 'B' for Dyne to Kilogram-Force")
            Console.WriteLine()
            Console.WriteLine("Enter R to return to the main menu")
            Console.WriteLine("Enter 'X' to Exit")
            forceInput = Console.ReadLine()
            forceInput = forceInput.ToUpper

            'Loops until user inputs "R" or "X"
            'Prompts user for valid input when invalid input is entered
            If forceInput = "A" Then
                forceKilotoDyne()

            ElseIf forceInput = "B" Then
                forceDyneToKilo()

            ElseIf forceInput = "R" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()

            ElseIf forceInput = "X" Then
                End

            Else
                Console.WriteLine("Invalid input, Please enter A, B or X, Press Enter to try again")
                Console.ReadLine()

            End If
        Loop Until forceInput = "R" Or forceInput = "X"
    End Sub
    Sub forceKilotoDyne()
        'Kilogram-force to Dyne converter
        'users can enter a kilogram-force unit and have it returned in Dyne
        Dim forceInput As Integer   'Stores Kilogram-Force Input and allows for massive whole numbers
        Dim ans As String           'Stores User response when prompted to convert again
        Dim validForce As Boolean   'Variable used in error catching process to avoid program crashing

        'Loop that will loop while ans = "Y"
        Do
            Console.Clear()
            Do      'Loop that will catch input errors and prompt user to input the required input
                Console.WriteLine("Enter the amount of Kilogram-Force you wish to convert: ")
                Try
                    validForce = True
                    forceInput = Console.ReadLine
                    Console.WriteLine("{0} Kilo's is equal to {1} Dyne", forceInput, forceToKiloDyne(forceInput))

                Catch
                    validForce = False
                    Console.WriteLine("Please Enter a valid figure")
                End Try
            Loop Until validForce = True

            Console.WriteLine()
            'Loop that prompts user to convert again and will loop until input is not "Y"
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")
            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            End If
        Loop While (ans = "Y")
    End Sub
    'The formula that carries out the Kilogram-force to Dyne conversion
    Function forceToKiloDyne(ByVal forceInput As Short) 'Takes a cop of the force figures entered by user

        Dim convertedKiloToDyne As Integer      'Stores Converted figures allowing for a wide range of whole numbers

        convertedKiloToDyne = forceInput * 980665       'Formula for converting Kilogram-Force into Dyne

        Return (convertedKiloToDyne)        'Returns Converted figures to calling code as Dyne

    End Function
    'Dyne to Kilogram force converter
    'User can enter a Dyne unit and have it returned in kilogram-force
    Sub forceDyneToKilo()
        Dim force As Integer        'Stores Dyne input entered by the user and allows for a wide range of whole numbers
        Dim ans As String           'Stores user input when prompted to convert again
        Dim validForce As Boolean   'Variable used in the error catching process to avoid program crashing

        Do      'Loop that will keep looping while input = "Y"
            Console.Clear()
            Do      'Loop that will catch input errors and prompt user to input the required input
                Console.WriteLine("Enter the amount of Dyne you wish to convert ")
                Try
                    validForce = True
                    force = Console.ReadLine
                    Console.WriteLine("{0} dyne is equal to {1} kilo", force, forceDyneToKilo(force))

                Catch
                    validForce = False
                    Console.WriteLine("Please enter a valid figure")
                End Try
            Loop Until validForce = True
            Console.WriteLine()

            'Loop that prompts user to convert again and will loop until input is not "Y"
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")
            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            End If
        Loop While (ans = "Y")

    End Sub
    'Formula that carries out the Dyne to kilogram-force conversion
    Function forceDyneToKilo(ByVal force As Short)      'takes a copy of the users input without affecting original memory location

        Dim DyneToKilo As Decimal      'Stores converted figures and allows for decimal numbers

        DyneToKilo = 0.000001019716213 * force      'Formula that converts Dyne into Kilogram-Force

        Return (DyneToKilo)     'Returns converted figures as Kilo to calling code
    End Function
    'Pressure Conversion menu
    'From here users can choose what pressure units to convert
    Sub pressureConversion()
        Dim pressureInput As String     'Stores user input at this menu and calls up appropriate Sub Procedures

        'Loop which will be activated by invalid input
        'Loop will end when input is "R" or "X"
        Do
            Console.Clear()
            Console.WriteLine("You have chosen Pressure Conversion")
            Console.WriteLine()
            Console.WriteLine("How would you like to convert?")
            Console.WriteLine("Enter 'A' for Bar to Dekapascal")
            Console.WriteLine("Enter 'B' for Dekapascal to Bar")
            Console.WriteLine()
            Console.WriteLine("Press R to return to the main menu")
            Console.WriteLine("Press X to Exit")
            pressureInput = Console.ReadLine()
            pressureInput = pressureInput.ToUpper
            If pressureInput = "A" Then
                pressureBarToDek()

            ElseIf pressureInput = "B" Then
                pressureDekToBar()

            ElseIf pressureInput = "R" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                'Console.ReadLine()

            ElseIf pressureInput = "X" Then
                Console.WriteLine("Program closing, Press Enter to continue")
                Console.ReadLine()
                End

            Else
                Console.WriteLine("Invalid input, Please enter A, B or X, Press Enter to continue")
                Console.ReadLine()
            End If
        Loop While pressureInput = "R" Or pressureInput = "X"
    End Sub
    'Bar to Dekapascal converter
    'User can enter Bar figures and have it returned in Dekspascal units
    Sub pressureBarToDek()
        Dim pressureInput As Integer        'Stores user input for conversion, allows a wide range of whole numbers
        Dim ans As String                   'Stores user input when prompted to convert again
        Dim validPressure As Boolean        'Variable that is used in error catching process to avoid program crashing

        Do
            Console.Clear()

            'Loop that will filter out errors and prompt user to input valid figures
            'When correct input is entered, Console will output converted figures
            Do
                Console.WriteLine("Enter the amount of Bar you wish to convert: ")
                Try
                    validPressure = True
                    pressureInput = Console.ReadLine()
                    Console.WriteLine()
                    Console.WriteLine("{0} Bar is equal to {1} Dekapascal", pressureInput, pressureBarToDek(pressureInput))
                    Console.WriteLine()
                Catch
                    validPressure = False
                    Console.WriteLine("Please enter a valid figure")
                End Try
            Loop Until validPressure = True

            'Prompts user to convert again
            'Will loop while answer is not "Y"
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")
            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            End If
        Loop While (ans = "Y")

    End Sub
    'Forumla that carries out the Bar to Dekapascal conversion
    Function pressureBarToDek(ByVal pressureInput As Short)     'Takes a copy of the users input without affecting the original memory location
        Dim convertedBarDek As Integer                  'Stores converted figures and allows for a wide variety of whole numbers

        convertedBarDek = pressureInput * 10000         'Formula that will convert Bar into Dekapascal
        Return (convertedBarDek)                        'Returns converted figures to calling code

    End Function
    'Dekapascal to Bar converter
    'User can enter Dekapascal units and have them returned in Bar
    Sub pressureDekToBar()
        Dim pressure As Integer     'Stores user input and allows for a wide variety of Whole Numbers
        Dim ans As String           'Stores User input when prompted to convert again
        Dim validPressure As Boolean    'Variable used in the error catching process to avoid program crashing

        'Loop that will filter out errors and prompt user to input valid figures
        'When correct input is entered, Console will output converted figures
        Do
            Console.Clear()

            Do
                Console.WriteLine("Enter the amount of Dekapascal you wish to convert: ")
                Try
                    validPressure = True
                    pressure = Console.ReadLine()
                    Console.WriteLine()
                    Console.WriteLine("{0} Dekapascal is equal to {1} Bar", pressure, dekToBar(pressure))
                Catch
                    validPressure = False
                    Console.WriteLine("Please enter a valid figure ")
                End Try
            Loop Until validPressure = True
            Console.WriteLine()

            'Prompts user to convert again
            'Will loop while answer is not "Y"
            Do
                Console.WriteLine("Would you like to convert again? [Y/N]")
                ans = Console.ReadLine()
                ans = ans.ToUpper
            Loop Until (ans = "Y" Or ans = "N")
            If ans = "N" Then
                Console.WriteLine("Returning to Main Menu, Press Enter to Continue")
                Console.ReadLine()
            End If
        Loop While (ans = "Y")


    End Sub
    'Formula that carries out the Dekapascal to Bar conversion
    Function dekToBar(ByVal pressure As Short)      'Takes a copy of the user's pressure input without effecting the original memory location
        Dim convertedDekBar As Decimal              'Stores converted figures and allows for Decimal numbers
        convertedDekBar = pressure * 0.0001         'Formula that will convert Dekapascal into Bar
        Return (convertedDekBar)                    'Returns converted figures in Bar to calling code
    End Function
End Module



