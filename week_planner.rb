#!/usr/bin/env ruby

require 'bundler/inline'

gemfile do
  source 'https://rubygems.org'
  gem 'calendav'
  gem 'icalendar'
end

require 'yaml'
require 'date'
require 'time'

# Get FastMail credentials from environment
FASTMAIL_EMAIL = ENV['FASTMAIL_EMAIL']
FASTMAIL_PASSWORD = ENV['FASTMAIL_APP_PASSWORD']
CALDAV_URL = ENV['CALDAV_URL']

unless FASTMAIL_EMAIL && FASTMAIL_PASSWORD && CALDAV_URL
  puts "Error: Missing credentials"
  puts ""
  puts "Please set the following environment variables:"
  puts "  export FASTMAIL_EMAIL='your@email.com'"
  puts "  export FASTMAIL_APP_PASSWORD='your-app-specific-password'"
  puts "  export CALDAV_URL='https://caldav.fastmail.com/dav/calendars/user/...'"
  puts ""
  puts "To create an app-specific password:"
  puts "  1. Go to https://www.fastmail.com/settings/security/devicekeys/new"
  puts "  2. Create a new app password with CalDAV access"
  puts "  3. Copy the generated password"
  exit 1
end

def next_monday
  today = Date.today
  days_until_monday = (1 - today.wday) % 7
  days_until_monday = 7 if days_until_monday == 0 # If today is Monday, get next Monday
  today + days_until_monday
end

def validate_monday(date_str)
  date = Date.parse(date_str)
  unless date.monday?
    puts "Error: #{date_str} is not a Monday"
    exit 1
  end
  date
rescue ArgumentError
  puts "Error: Invalid date format '#{date_str}'. Use YYYY-MM-DD"
  exit 1
end

# Parse command line arguments
debug_mode = ARGV.include?('--debug')
date_arg = ARGV.find { |arg| !arg.start_with?('--') }

week_start = if date_arg
               validate_monday(date_arg)
             else
               next_monday
             end

week_end = week_start + 6

# Load budget
budget = YAML.load_file('budget.yml')['categories']

# Validate budget totals 168 hours
total_budgeted = budget.values.sum
unless total_budgeted == 168
  puts "Error: Budget does not add up to 168 hours (7 days × 24 hours)"
  puts "Current total: #{total_budgeted} hours"
  puts "Difference: #{(total_budgeted - 168).round(1)} hours"
  exit 1
end

# Fetch calendar events
credentials = Calendav.credentials(
  :fastmail,
  FASTMAIL_EMAIL,
  FASTMAIL_PASSWORD
)

# Monkey-patch OpenSSL to disable SSL verification for this script
# (reading calendar data only, low security risk)
module OpenSSL
  module SSL
    class SSLContext
      alias_method :original_set_params, :set_params

      def set_params(params = {})
        original_set_params(params)
        self.verify_mode = OpenSSL::SSL::VERIFY_NONE
        self
      end
    end
  end
end

client = Calendav::Client.new(credentials)

# Query events for the week
# Convert to Time objects with start of day and end of day
time_min = Time.new(week_start.year, week_start.month, week_start.day, 0, 0, 0)
time_max = Time.new(week_end.year, week_end.month, week_end.day, 23, 59, 59)

calendar_objects = client.events.list(
  CALDAV_URL,
  from: time_min,
  to: time_max
)

# Debug mode: print all events and exit
if debug_mode
  puts "DEBUG: All calendar events for #{week_start} to #{week_end}"
  puts "=" * 80
  puts "CalDAV URL: #{CALDAV_URL}"
  puts "Time range: #{time_min} to #{time_max}"
  puts ""
  puts "Number of calendar objects returned: #{calendar_objects.size}"
  puts ""

  if calendar_objects.empty?
    puts "No calendar objects returned from CalDAV server."
    puts ""
    puts "Possible issues:"
    puts "- Wrong calendar URL"
    puts "- Authentication failed silently"
    puts "- No events in this time range"
    puts "- Calendar permissions issue"
  else
    puts "Parsed events:"
    puts ""

    calendar_objects.each_with_index do |cal_object, idx|
      calendar = Icalendar::Calendar.parse(cal_object.calendar_data).first
      next unless calendar

      calendar.events.each do |event|
        title = event.summary.to_s.strip
        start_time = event.dtstart

        # Calculate duration - handle both DTEND and DURATION formats
        if event.dtend
          end_time = event.dtend
          duration_seconds = end_time.to_time - start_time.to_time
        elsif event.duration
          # Duration is an OpenStruct with hours, minutes, seconds properties
          dur = event.duration
          duration_seconds = (dur.weeks * 604800) + (dur.days * 86400) +
                           (dur.hours * 3600) + (dur.minutes * 60) + dur.seconds
          end_time = start_time.to_time + duration_seconds
        else
          # Skip events with no end time or duration
          next
        end

        duration_hours = duration_seconds / 3600.0

        puts "Event #{idx + 1}:"
        puts "  Title: #{title}"
        puts "  Start: #{start_time}"
        puts "  End: #{end_time}"
        puts "  Duration: #{duration_hours.round(2)} hours"

        # Check for recurrence rules and expand
        if event.rrule && !event.rrule.empty?
          rrule = event.rrule.first
          puts "  RRULE: #{rrule.frequency} on #{rrule.by_day&.join(', ')}"

          if rrule.frequency == "WEEKLY" && rrule.by_day
            day_map = { "SU" => 0, "MO" => 1, "TU" => 2, "WE" => 3, "TH" => 4, "FR" => 5, "SA" => 6 }
            target_wdays = rrule.by_day.map { |d| day_map[d] }

            matching_dates = []
            (week_start..week_end).each do |date|
              matching_dates << date if target_wdays.include?(date.wday)
            end

            puts "  Expands to #{matching_dates.size} occurrences in week:"
            matching_dates.each { |d| puts "    - #{d}" }
            puts "  Total duration for week: #{(duration_hours * matching_dates.size).round(2)} hours"
          end
        end

        # Show which budget category it would match
        matched_category = budget.keys.find { |category| title.start_with?(category) }
        if matched_category
          puts "  Matches budget category: #{matched_category}"
        else
          puts "  NO MATCH - would be categorized as Uncategorized"
        end
        puts ""
      end
    end
  end
  exit 0
end

# Parse iCalendar data and aggregate hours
actual_hours = Hash.new(0.0)
uncategorized_events = []

calendar_objects.each do |cal_object|
  calendar = Icalendar::Calendar.parse(cal_object.calendar_data).first
  next unless calendar

  calendar.events.each do |event|
    title = event.summary.to_s.strip
    start_time = event.dtstart

    # Calculate duration - handle both DTEND and DURATION formats
    if event.dtend
      end_time = event.dtend
      duration_seconds = end_time.to_time - start_time.to_time
    elsif event.duration
      # Duration is an OpenStruct with weeks, days, hours, minutes, seconds properties
      dur = event.duration
      duration_seconds = (dur.weeks * 604800) + (dur.days * 86400) +
                       (dur.hours * 3600) + (dur.minutes * 60) + dur.seconds
    else
      # Skip events with no end time or duration
      next
    end

    duration_hours = duration_seconds / 3600.0

    # Handle recurring events
    occurrences = []
    if event.rrule && !event.rrule.empty?
      # Expand recurrence within the week
      rrule = event.rrule.first

      if rrule.frequency == "WEEKLY" && rrule.by_day
        # Map day abbreviations to Ruby wday numbers
        day_map = { "SU" => 0, "MO" => 1, "TU" => 2, "WE" => 3, "TH" => 4, "FR" => 5, "SA" => 6 }
        target_wdays = rrule.by_day.map { |d| day_map[d] }

        # Find all matching days in the week
        (week_start..week_end).each do |date|
          if target_wdays.include?(date.wday)
            occurrences << duration_hours
          end
        end
      else
        # For other recurrence types, just count the base occurrence
        occurrences << duration_hours
      end
    else
      # Non-recurring event
      occurrences << duration_hours
    end

    # Find matching budget category (event title starts with category name)
    matched_category = budget.keys.find { |category| title.start_with?(category) }

    total_hours = occurrences.sum
    if matched_category
      actual_hours[matched_category] += total_hours
    else
      uncategorized_events << { title: title, hours: total_hours }
      actual_hours["Uncategorized"] += total_hours
    end
  end
end

# Build comparison table data
table_data = []

budget.each do |category, budgeted|
  actual = actual_hours[category] || 0.0
  variance = actual - budgeted

  table_data << {
    category: category,
    budgeted: budgeted.round(1),
    actual: actual.round(1),
    variance: variance.round(1)
  }
end

# Add uncategorized row if there are uncategorized events
if actual_hours["Uncategorized"] > 0
  table_data << {
    category: "Uncategorized",
    budgeted: 0.0,
    actual: actual_hours["Uncategorized"].round(1),
    variance: actual_hours["Uncategorized"].round(1)
  }
end

# Add totals row
total_budgeted = budget.values.sum
total_actual = actual_hours.values.sum
table_data << {
  category: "TOTAL",
  budgeted: total_budgeted.round(1),
  actual: total_actual.round(1),
  variance: (total_actual - total_budgeted).round(1)
}

# Output using gum table
# gum table expects CSV with proper quoting for fields containing special characters
puts "Category,Budgeted,Actual,Variance"
table_data.each do |row|
  # Quote category names that contain colons or commas
  category = row[:category].include?(',') || row[:category].include?(':') ? "\"#{row[:category]}\"" : row[:category]
  puts "#{category},#{row[:budgeted]},#{row[:actual]},#{row[:variance]}"
end
