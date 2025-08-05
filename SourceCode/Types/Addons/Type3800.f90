Subroutine Type3800
   !------------------------------------------------------------------------------------
   !    DESCRIPTION
   !------------------------------------------------------------------------------------
   ! Subroutine Type3800 is a custom TRNSYS component that converts the simulation time
   ! into the Excel serial datetime format
   !
   ! This component enables compatibility with spreadsheet tools such as Microsoft Excel
   ! by converting TRNSYS simulation time into a format where dates and times
   ! are represented as serial numbers, with January 1, 1900 corresponding to 1.0
   !
   ! Parameters:
   !   1. Mode (1 or 2)
   !      - 1: Time accumulates over multiple years from a reference year
   !      - 2: Time loops within the reference year
   !   2. Reference Year: Starting year for the conversion (e.g., 2025)
   !
   ! Outputs:
   !   - Excel Serial Date: A real number compatible with Excel datetime format
   !
   ! Notes::
   !   - TRNSYS simulations always assume a non-leap year, with 8760 hours per year
   !   - As a result, February 29th is never simulated and will be skipped in the conversion
   !   - To ensure a continuous date mapping to Excel serial numbers,
   !     it is recommended to avoid using leap years as the reference year
   !
   ! Version History:
   !   2025-05-01 – A. Lachance: Original implementation
   !
   !DEC$ATTRIBUTES DLLEXPORT :: TYPE3800
   !------------------------------------------------------------------------------------
   !    VARIABLES
   !------------------------------------------------------------------------------------
   Use TrnsysConstants
   Use TrnsysFunctions
   Implicit None

   Double Precision :: timestep, time
   Integer :: currentUnit, currentType

   integer :: reference_year
   integer :: mode
   integer :: simulation_year
   integer :: adjusted_year
   integer :: adjusted_month
   integer :: century_adjustment
   integer :: days_since_julian_origin
   integer :: days_to_excel_epoch
   integer :: day_of_year
   integer :: decimal_hours
   integer :: minute_value
   integer :: current_month
   integer :: current_day

   double precision :: elapsed_hours
   double precision :: hour_with_fraction
   double precision :: time_of_day_fraction
   double precision :: hours_in_day
   double precision :: excel_serial_date

   !------------------------------------------------------------------------------------
   !    INITIALIZATION
   !------------------------------------------------------------------------------------
   ! Get global TRNSYS simulation variables
   time = getSimulationTime()
   timestep = getSimulationTimeStep()
   currentUnit = getCurrentUnit()
   currentType = getCurrentType()

   !------------------------------------------------------------------------------------
   !    VERSION CHECK
   !------------------------------------------------------------------------------------
   If (getIsVersionSigningTime()) Then
      Call SetTypeVersion(17)
      Return
   End If

   !------------------------------------- ----------------------------------------------
   !    FINAL CALL HANDLING
   !------------------------------------------------------------------------------------
   If (getIsLastCallofSimulation() .or. getIsEndOfTimestep()) Then
      mode = getParameterValue(1)
      reference_year = getParameterValue(2)
      time = getSimulationTime()

      ! Determine the simulation year
      select case (mode)
      case (1)
         simulation_year = reference_year + int(time/8760.0d0)
      case (2)
         simulation_year = reference_year
      end select

      ! Compute day of year and fractional hour
      time = mod(time, 8760.0d0)
      day_of_year = int(time/24.0d0) + 1
      hours_in_day = mod(time, 24.0d0)
      decimal_hours = int(hours_in_day)
      minute_value = int((hours_in_day - decimal_hours)*60.0d0)
      hour_with_fraction = hours_in_day

      ! Convert day of year to month and day
      select case (day_of_year)
      case (:31)
         current_month = 1
         current_day = day_of_year
      case (32:59)
         current_month = 2
         current_day = day_of_year - 31
      case (60:90)
         current_month = 3
         current_day = day_of_year - 59
      case (91:120)
         current_month = 4
         current_day = day_of_year - 90
      case (121:151)
         current_month = 5
         current_day = day_of_year - 120
      case (152:181)
         current_month = 6
         current_day = day_of_year - 151
      case (182:212)
         current_month = 7
         current_day = day_of_year - 181
      case (213:243)
         current_month = 8
         current_day = day_of_year - 212
      case (244:273)
         current_month = 9
         current_day = day_of_year - 243
      case (274:304)
         current_month = 10
         current_day = day_of_year - 273
      case (305:334)
         current_month = 11
         current_day = day_of_year - 304
      case (335:365)
         current_month = 12
         current_day = day_of_year - 334
      end select

      ! Convert date to days since 0000-03-01 (Julian day formula)
      if (current_month <= 2) then
         adjusted_year = simulation_year - 1
         adjusted_month = current_month + 12
      else
         adjusted_year = simulation_year
         adjusted_month = current_month
      end if

      century_adjustment = adjusted_year/100
      days_since_julian_origin = 365*adjusted_year + adjusted_year/4 - century_adjustment + &
                                 century_adjustment/4 + int(30.6001*(adjusted_month + 1)) + &
                                 current_day - 679006

      ! Excel epoch: 1899-12-30
      adjusted_year = 1899
      adjusted_month = 12
      if (adjusted_month <= 2) then
         adjusted_year = adjusted_year - 1
         adjusted_month = adjusted_month + 12
      end if

      century_adjustment = adjusted_year/100
      days_to_excel_epoch = 365*adjusted_year + adjusted_year/4 - century_adjustment + &
                           century_adjustment/4 + int(30.6001*(adjusted_month + 1)) + 30 - 679006

      ! Time fraction of the day
      time_of_day_fraction = hour_with_fraction/24.0d0

      ! Excel serial number
      excel_serial_date = dble(days_since_julian_origin - days_to_excel_epoch) + dble(time_of_day_fraction)
      ! Output results
      Call setOutputValue(1, excel_serial_date)
      
      Return
   End If

   !------------------------------------------------------------------------------------
   !    TYPE INITIALIZATION
   !------------------------------------------------------------------------------------
   If (getIsFirstCallofSimulation()) Then
      mode = getParameterValue(1)
      reference_year = getParameterValue(2)

      If ((mode < 1) .or. (mode > 2)) Then
         Call FoundBadParameter(1, 'Fatal', ' must be between value of 1 and 2.')
      End If

      If ((reference_year <= 1900) .or. (reference_year > 3000)) Then
         Call FoundBadParameter(2, 'Fatal', ' must be between value of 1900 and 3000.')
      End If

      If (ErrorFound()) Return

      ! TRNSYS Engine Type Calls
      Call SetNumberofParameters(2)
      Call SetNumberofInputs(0)
      Call SetNumberofDerivatives(0)
      Call SetNumberofOutputs(1)
      Call SetIterationMode(3)
      Call SetNumberStoredVariables(0, 0)
      Call SetNumberofDiscreteControls(0)

      ! TRNSYS Input and Output Units
      ! Nothing

      Return
   End If

   !------------------------------------------------------------------------------------
   !    INITIAL VALUE SETTING
   !------------------------------------------------------------------------------------
   If (getIsStartTime()) Then
      Call setOutputValue(1, 0d0)
      Return
   End If

   !------------------------------------------------------------------------------------
   !    RE-READ PARAMETERS
   !------------------------------------------------------------------------------------
   If (getIsReReadParameters()) Then
      mode = getParameterValue(1)
      reference_year = getParameterValue(2)
   End If

   !------------------------------------------------------------------------------------
   !    MAIN TYPE CODE
   !------------------------------------------------------------------------------------

   Return
End
