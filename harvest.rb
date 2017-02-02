#!/bin/ruby

require 'twitter'
require 'pp'
require 'rubyXL'

client = Twitter::REST::Client.new do |config|
  config.consumer_key        = ENV['TWITTER_CONSUMER_KEY'] || nil
  config.consumer_secret     = ENV['TWITTER_CONSUMER_SECRET'] || nil
  config.access_token        = ENV['TWITTER_ACCESS_TOKEN'] || nil
  config.access_token_secret = ENV['TWITTER_ACCESS_TOKEN_SECRET'] || nil
end

workbook = RubyXL::Workbook.new

TWEET_HEADERS = {:user => 'Username',
                 :text => 'Text',
                 :created_at => 'Created At',
                 :url => 'Tweet URL',
                 :uri => 'Tweet URI',
                 :favorited? => 'Has Been Favorited By This Account?',
                 :favorite_count => 'Number of Times Tweet Favorited By Other Account(s)',
                 :geo => 'Geo Information',
                 :hashtags? => 'Contains Hashtags?',
                 :hashtags => 'Hashtags',
                 :id => 'Tweet ID',
                 :in_reply_to_screen_name => 'In Reply To Screen Name',
                 :in_reply_to_user_id => 'In Reply To User ID',
                 :in_reply_to_status_id => 'In Reply To Tweet ID',
                 :lang => 'Language',
                 :media? => 'Contains Media?',
                 :media => 'Media',
                 :metadata => 'Metadata',
                 :place => 'Place',
                 :possibly_sensitive? => 'Possibly Sensitive',
                 :quote? => 'Is a Quote?',
                 :quoted_status => 'Quoted Status',
                 :reply? => 'Is a Reply?',
                 :retweet_count => 'Number of Times Retweeted',
                 :retweet? => 'Is a Retweet?',
                 :retweeted_status => 'Retweeted Status',
                 :source => 'Source',
                 :symbols? => 'Contains Symbols?',
                 :symbols => 'Symbols',
                 :truncated? => 'Is Truncated?',
                 :urls? => 'Contains URLs?',
                 :urls => 'URLs',
                 :user_mentions => 'User Mentions',
                 :filter_level? => 'Has Filter Level?',
                 :filter_level => 'Filter Level'}.freeze

TWEET_EXPLODERS = [:user, :urls, :user_mentions, :hashtags, :media, :quoted_status, :retweeted_status].freeze

TWEET_STRINGIFY_FIELDS = [:id,
                          :in_reply_to_screen_name,
                          :in_reply_to_user_id,
                          :in_reply_to_status_id,
                          :in_reply_to_tweet_id].freeze

USER_HEADERS = [:id, :name, :screen_name, :description,
                :location, :url, :contributors_enabled, :created_at,
                :default_profile_image, :default_profile, :favourites_count, :followers_count,
                :friends_count, :geo_enabled, :has_extended_profile, :is_translation_enabled,
                :is_translator, :lang, :listed_count, :profile_background_color,
                :profile_background_image_url_https, :profile_background_image_url, :profile_background_tile, :profile_banner_url,
                :profile_image_url_https, :profile_image_url, :profile_link_color, :profile_location,
                :profile_sidebar_border_color, :profile_sidebar_fill_color, :profile_text_color, :profile_use_background_image,
                :protected, :statuses_count, :time_zone, :translator_type,
                :utc_offset, :verified].freeze

def client.load_and_validate_arguments(arguments)
  abort('Specify a valid, public Twitter username (case-sensitive) and an integer from 0 - 200.') if arguments.length < 2
  user = arguments.first
  count = arguments.last.to_i
  begin
    user(user)
  rescue
    abort("Invalid Twitter username #{user} specified.")
  end
  return user, count
end

def client.record_tweets(user, count, workbook)
  iterator = count/200 > 0 ? count/200 : 1
  remainder = count > 200 ? count%200 : 0
  iterator += 1 unless remainder == 0
  max_id = nil
  y_index = 1
  rate_limit_hit = false
  (1..iterator).each do |i|
    num = (iterator == i && remainder > 0) ? remainder : (count < 200 ? count : 200)
    options = {count: num, include_rts: true}
    options[:max_id] = max_id unless max_id.nil?
    tweets = user_timeline(user, options)
    tweets.each_with_index do|tweet, index|
      explode(workbook[0], tweet, y_index)
      y_index+=1
    end
    rate_limit_hit = true if tweets.empty? && count >= 3200
    max_id = tweets.last.id - 1 unless tweets.empty?
  end
  rate_limit_hit ? "Maximum tweets allowed or available (#{y_index}) harvesting from #{user}." : "Tweets harvested from #{user}."
end

def workbook.set_up_spreadsheet(user)
  worksheet = worksheets[0]
  worksheet.sheet_name = 'tweets'
  add_user_worksheet(user)
  set_headers(worksheet, TWEET_HEADERS)
end

def set_headers(worksheet, headers_hash)
  headers_hash.each_with_index do |(key, value), index|
    worksheet.add_cell(0,index, value)
  end
end

def explode(worksheet, tweet, y_index)
  TWEET_HEADERS.each_with_index do |(key, value), x_index|
    worksheet.add_cell(y_index, x_index, TWEET_EXPLODERS.member?(key) ? sanitized(exploded_value(tweet, key)) : sanitized(typed_value(tweet, key)))
  end
end

def exploded_value(tweet, key)
  return tweet.user.name if key == :user
  return pp(tweet.attrs[:entities][:media]) if key == :media
  return pp(tweet.attrs[:entities][:urls]) if key == :urls
  return pp_exploded_retweet(tweet) if key == :quoted_status && tweet.quote?
  return pp_exploded_retweet(tweet) if key == :retweeted_status && tweet.retweet?
  return pp_array(tweet.hashtags) if key == :hashtags
  return pp_attrs_hash(tweet.user_mentions) if key == :user_mentions
end

def typed_value(tweet, key)
  return TWEET_STRINGIFY_FIELDS.member?(key) ? tweet.send(key).to_s : tweet.send(key)
end

def sanitized(value)
  return value.to_s == '[]' ? '' : value
end

def workbook.add_user_worksheet(user)
  add_worksheet('user info')
  worksheet = self[1]
  headers_hash = {}
  USER_HEADERS.each{|header| headers_hash[header] = header_label(header) }
  set_headers(worksheet, headers_hash)
  USER_HEADERS.each_with_index do |key, index|
    worksheet.add_cell(1, index, user.attrs[key].to_s)
  end
end

def pp_array(array)
  p_array = []
  array.each{|a| p_array.push(a.text)}
  return p_array.empty? ? nil : p_array
end

def pp_attrs_hash(attrs_hash)
  pp_string = ''
  attrs_hash.each do |entry|
    pp_string += "#{pp(entry.attrs)}"
  end
  return pp_string.empty? ? pp_string : nil
end



def pp_exploded_retweet(tweet)
  return pp(tweet.attrs[:entities][:urls]).to_s
end

def header_label(header_symbol)
  header_symbol.to_s.capitalize.gsub('_',' ')
end

begin
  client.verify_credentials
  user, count = client.load_and_validate_arguments(ARGV)
rescue Twitter::Error::Forbidden => exception
  abort("#{exception.message}. Make sure TWITTER_CONSUMER_KEY, TWITTER_CONSUMER_SECRET, TWITTER_ACCESS_TOKEN, and TWITTER_ACCESS_TOKEN_SECRET are set as environment variables.")
end


begin
  workbook.set_up_spreadsheet(client.user(user))
  message = client.record_tweets(user, count, workbook)
  spreadsheet_name = "#{user}_#{DateTime.now.to_time.to_i}.xlsx"
  puts 'Writing spreadsheet...'
  workbook.write(spreadsheet_name)
  puts "#{message} Spreadsheet written to #{spreadsheet_name}."
rescue Twitter::Error::TooManyRequests => error
  puts 'Rate limit hit, sleeping...'
  sleep error.rate_limit.reset_in + 1
  retry
rescue Twitter::Error::InternalServerError
  puts 'Twitter Internal Server Error, retrying harvest...'
  retry
rescue Twitter::Error::ServiceUnavailable
  abort('Twitter service unavailable.')
rescue Twitter::Error::Unauthorized
  abort('This account is protected, unable to harvest.')
end
