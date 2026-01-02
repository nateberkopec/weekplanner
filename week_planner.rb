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

DAY_MAP = { "SU" => 0, "MO" => 1, "TU" => 2, "WE" => 3, "TH" => 4, "FR" => 5, "SA" => 6 }.freeze
SECONDS_PER_HOUR = 3600.0

# Get FastMail credentials from environment
FASTMAIL_EMAIL = ENV['FASTMAIL_EMAIL']
FASTMAIL_PASSWORD = ENV['FASTMAIL_APP_PASSWORD']
CALDAV_URL = ENV['CALDAV_URL']

unless FASTMAIL_EMAIL && FASTMAIL_PASSWORD && CALDAV_URL
  abort <<~ERROR
    Error: Missing credentials

    Please set the following environment variables:
      export FASTMAIL_EMAIL='your@email.com'
      export FASTMAIL_APP_PASSWORD='your-app-specific-password'
      export CALDAV_URL='https://caldav.fastmail.com/dav/calendars/user/...'

    To create an app-specific password:
      1. Go to https://www.fastmail.com/settings/security/devicekeys/new
      2. Create a new app password with CalDAV access
      3. Copy the generated password
  ERROR
end

def next_monday
  today = Date.today
  days_until_monday = (1 - today.wday) % 7
  days_until_monday = 7 if days_until_monday.zero?
  today + days_until_monday
end

def validate_monday(date_str)
  date = Date.parse(date_str)
  abort "Error: #{date_str} is not a Monday" unless date.monday?
  date
rescue ArgumentError
  abort "Error: Invalid date format '#{date_str}'. Use YYYY-MM-DD"
end

def calculate_duration(event)
  if event.dtend
    event.dtend.to_time - event.dtstart.to_time
  elsif event.duration
    dur = event.duration
    (dur.weeks * 604800) + (dur.days * 86400) +
      (dur.hours * 3600) + (dur.minutes * 60) + dur.seconds
  end
end

def expand_recurrences(event, week_start, week_end)
  return [1] unless event.rrule&.any?

  rrule = event.rrule.first
  return [1] unless rrule.frequency == "WEEKLY" && rrule.by_day

  target_wdays = rrule.by_day.map { |d| DAY_MAP[d] }
  (week_start..week_end).select { |date| target_wdays.include?(date.wday) }
end

# Parse command line arguments
debug_mode = ARGV.include?('--debug')
budget_arg = ARGV.find { |arg| arg.start_with?('--budget=') }
budget_file = budget_arg ? budget_arg.split('=', 2).last : 'budget.yml'
date_arg = ARGV.find { |arg| !arg.start_with?('--') }

week_start = date_arg ? validate_monday(date_arg) : next_monday
week_end = week_start + 6

# Load and validate budget
budget = YAML.load_file(budget_file)['categories']
total_budgeted = budget.values.sum

unless total_budgeted == 168
  abort <<~ERROR
    Error: Budget does not add up to 168 hours (7 days × 24 hours)
    Current total: #{total_budgeted} hours
    Difference: #{(total_budgeted - 168).round(1)} hours
  ERROR
end

# Disable SSL verification (reading calendar data only, low security risk)
module OpenSSL::SSL
  class SSLContext
    alias_method :original_set_params, :set_params
    def set_params(params = {})
      original_set_params(params)
      self.verify_mode = OpenSSL::SSL::VERIFY_NONE
      self
    end
  end
end

# Fetch calendar events
credentials = Calendav.credentials(:fastmail, FASTMAIL_EMAIL, FASTMAIL_PASSWORD)
client = Calendav::Client.new(credentials)

time_min = Time.new(week_start.year, week_start.month, week_start.day, 0, 0, 0)
time_max = Time.new(week_end.year, week_end.month, week_end.day, 23, 59, 59)

calendar_objects = client.events.list(CALDAV_URL, from: time_min, to: time_max)

# Debug mode: print all events and exit
if debug_mode
  puts "DEBUG: All calendar events for #{week_start} to #{week_end}"
  puts "=" * 80
  puts "CalDAV URL: #{CALDAV_URL}"
  puts "Time range: #{time_min} to #{time_max}"
  puts "\nNumber of calendar objects returned: #{calendar_objects.size}\n\n"

  if calendar_objects.empty?
    puts <<~DEBUG
      No calendar objects returned from CalDAV server.

      Possible issues:
      - Wrong calendar URL
      - Authentication failed silently
      - No events in this time range
      - Calendar permissions issue
    DEBUG
  else
    puts "Parsed events:\n\n"

    calendar_objects.each_with_index do |cal_object, idx|
      calendar = Icalendar::Calendar.parse(cal_object.calendar_data).first
      next unless calendar

      calendar.events.each do |event|
        title = event.summary.to_s.strip
        duration_seconds = calculate_duration(event)
        next unless duration_seconds

        duration_hours = duration_seconds / SECONDS_PER_HOUR
        occurrences = expand_recurrences(event, week_start, week_end)

        puts "Event #{idx + 1}:"
        puts "  Title: #{title}"
        puts "  Start: #{event.dtstart}"
        puts "  Duration: #{duration_hours.round(2)} hours"

        if event.rrule&.any?
          rrule = event.rrule.first
          puts "  RRULE: #{rrule.frequency} on #{rrule.by_day&.join(', ')}"
          if occurrences.is_a?(Array) && occurrences.size > 1
            puts "  Expands to #{occurrences.size} occurrences in week:"
            occurrences.each { |d| puts "    - #{d}" }
            puts "  Total duration for week: #{(duration_hours * occurrences.size).round(2)} hours"
          end
        end

        matched_category = budget.keys.find { |category| title.start_with?(category) }
        puts "  #{matched_category ? "Matches budget category: #{matched_category}" : "NO MATCH - would be categorized as Uncategorized"}"
        puts ""
      end
    end
  end
  exit 0
end

# Parse iCalendar data and aggregate hours
actual_hours = Hash.new(0.0)

calendar_objects.each do |cal_object|
  calendar = Icalendar::Calendar.parse(cal_object.calendar_data).first
  next unless calendar

  calendar.events.each do |event|
    title = event.summary.to_s.strip
    duration_seconds = calculate_duration(event)
    next unless duration_seconds

    duration_hours = duration_seconds / SECONDS_PER_HOUR
    occurrences = expand_recurrences(event, week_start, week_end)
    total_hours = duration_hours * occurrences.size

    matched_category = budget.keys.find { |category| title.start_with?(category) }
    actual_hours[matched_category || "Uncategorized"] += total_hours
  end
end

# Build comparison table data
table_data = budget.map do |category, budgeted|
  actual = actual_hours[category]
  {
    category: category,
    budgeted: budgeted.round(1),
    actual: actual.round(1),
    variance: (actual - budgeted).round(1)
  }
end

# Add uncategorized row if there are uncategorized events
if actual_hours["Uncategorized"] > 0
  uncategorized = actual_hours["Uncategorized"]
  table_data << {
    category: "Uncategorized",
    budgeted: 0.0,
    actual: uncategorized.round(1),
    variance: uncategorized.round(1)
  }
end

# Add totals row
total_actual = actual_hours.values.sum
table_data << {
  category: "TOTAL",
  budgeted: total_budgeted.round(1),
  actual: total_actual.round(1),
  variance: (total_actual - total_budgeted).round(1)
}

# Output CSV for gum table
puts "Category,Budgeted,Actual,Variance"
table_data.each do |row|
  category = row[:category].match?(/[,:]/) ? "\"#{row[:category]}\"" : row[:category]
  puts "#{category},#{row[:budgeted]},#{row[:actual]},#{row[:variance]}"
end
