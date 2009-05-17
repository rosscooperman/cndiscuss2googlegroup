#!/usr/bin/env ruby

# require some built-in ruby libraries
require 'optparse'
require 'cgi'
require 'net/smtp'

# and some gems
require 'rubygems'
require 'mechanize'
require 'highline/import'

# reopen the string class and add some header finding methods
class String
  def mail_header(header)
    matches = match /^#{header}: \s*(.+)\s*$/i
    CGI::unescapeHTML(matches[1]) if matches
  end
  
  def formatted_mail_header(header)
    "#{header.capitalize}: #{mail_header(header)}"
  end
  
  def extract_headers
    # strip leading (and trailing) <td><pre> tags and whitespace and unescape HTML entities
    cn_header = gsub /^\s*<td><pre>\s*|\s*<\/pre><\/td>\s*$/, ''

    {
      :to           => cn_header.mail_header('to'),
      :from         => cn_header.mail_header('from'),
      :subject      => cn_header.mail_header('subject'),
      :date         => cn_header.mail_header('date'),
      :content_type => "text/html; charset=UTF-8",
      :x_mailer     => "CollabNet to Google Groups v1.0"
    }
  end
  
  def shorten(length)
    if self.length > length - 3
      self[0...(length - 3)] + '...'
    else
      self
    end
  end
end

# reopen the Hash class to add some to_mail options
class Hash
  # convert a hash to mail headers
  def to_mailh
    self.inject('') do |str, header|
      str += "#{header[0].to_s.gsub(/_/, ' ').capitalize.gsub(/ /, '-')}: #{header[1]}\n"
    end + "\n"
  end
  
  # convert a hash to mail message
  def to_mail
    if key?(:headers) && key?(:body)
      self[:headers].to_mailh + self[:body]
    end
  end
end

# open up some Mechanize classes to add some utility methods
module WWW
  class Mechanize
    class Page
      def filtered_links(filter)
        links.select { |l| l.href =~ /^#{filter}/ }
      end
    end
  end
end

# main class of this application
class App
  VERSION = '1.0'
  COMMAND = File.basename(__FILE__)

  # initializer method
  def initialize
    # default options
    @options = {
      :username => ENV['USER'] || ENV['USERNAME'],
      :skip     => 0,
      :limit    => -1,
      :strip    => ''
    }
    
    # Mechanize agent to use for fetching web pages
    @agent = WWW::Mechanize.new { |agent|
      agent.user_agent_alias = 'Mac Safari'
      agent.keep_alive       = false
    }
  end
  
  # parse command line options and required parameters
  def options_parse
    # first, deal with all options
    opts = OptionParser.new do |opts|
      opts.banner  = "Usage: #{COMMAND} [options] <cn_base_url>\n"
      opts.banner += "Convert CollabNet discussions into mail messages"

      opts.on("-h", "Show this message") do
        puts opts
        exit
      end

      opts.on("-v", "Display this application's version and exit") do
        puts "#{COMMAND} version #{VERSION}"
        exit
      end
      
      opts.on("-s <number>", "Skip <number> messages before processing") do |number|
        @options[:skip] = number.to_i
      end

      opts.on("-l <number>", "Limit conversion to <number> messages") do |number|
        @options[:limit] = number.to_i
      end

      opts.on("-u <username>", "Username for logging in to CollabNet") do |username|
        @options[:username] = username
      end
      
      opts.on("-f <forum>", "Forum that is being converted") do |forum|
        @options[:forum] = forum
      end
      
      opts.on("-e <address>", "Send mail to <address> rather than ignoring") do |address|
        @options[:address] = address
      end
      
      opts.on("-t <string>", "Strip <string> from the start of subject lines") do |string|
        @options[:strip] = string
      end
      
      opts.on("-d <min>:<max>", "Delay between <min> and <max> seconds between messages") do |val|
        unless val =~ /^(\d+):(\d+)$/
          puts "invalid delay range -- #{val}\n\n#{opts}"
          exit
        end
        @options[:min] = $1.to_i
        @options[:max] = $2.to_i
      end
      
    end
    
    # try parsing options, rescuing certain errors and printing friendly messages
    begin
      opts.parse!
    rescue OptionParser::InvalidOption => e
      puts "#{e}\n\n#{opts}"
      exit
    end
    
    # now, make sure that there are enough arguments left for required parameters
    if ARGV.length < 1
      puts "You must specify a CollabNet base URL (<cn_base_url>)\n\n#{opts}"
      exit
    else
      @options[:base_url] = ARGV[0]
    end
    
    generate_urls
  end
    
  # "main" method for this application
  def run
    options_parse

    # get the user's password but echo * instead of showing the password
    @options[:password] = ask("Password for '#{@options[:username]}': ") { |q| q.echo = '*' }
    
    # log into the CollabNet instance
    login
    
    forums.each do |name, url|
      puts "== Forum #{name}: extracting ".ljust(80, '=')
      start = Time.now
      
      # fetch messages for this forum
      messages(url).each do |url|
        message_start = Time.now
        # check for limit/skip conditions
        break if @options[:limit] == 0
                  
        message = message_to_mail(url)
        
        if @options[:skip] > 0
          @options[:skip] -= 1
          puts "-- Skipped message (#{message[:headers][:subject]}".shorten(69) + ')'
          puts "   -> #{sprintf("%.4f", Time.now - message_start)}s"
          next
        elsif !@options[:address].nil?
          Net::SMTP.start('localhost') do |smtp|
            smtp.send_message(message.to_mail, message[:headers][:from], @options[:address])
          end
          puts "-- Sent message (#{message[:headers][:subject]}".shorten(69) + ')'
        else
          puts "-- Ignored message (#{message[:headers][:subject]}".shorten(69) + ')'
        end

        puts "   -> #{sprintf("%.4f", Time.now - message_start)}s"
        @options[:limit] -= 1
        
        if @options[:max]
          sleep_time = rand(@options[:max] - @options[:min] + 1) + @options[:min]
          puts "-- Delaying (#{sleep_time}s)"
          sleep(sleep_time)
        end
      end
            
      delta = sprintf("%.4f", Time.now - start)
      puts "== Forum #{name}: extracted (#{(delta)}s) ".ljust(80, '=') + "\n"
      break if @options[:limit] == 0
    end
  end
  
  private
  
  # generate URLs based on options chosen 
  def generate_urls
    # strip trailing / from the base URL
    base_url = @options[:base_url].gsub /\/$/, ''
    
    # set up a hash of URLs
    @urls = {
      :login  => base_url.gsub(/^https?:\/\/\w+\.(.+)$/, 'http://www.\1') + "/servlets/Login",
      :forums => "#{base_url}/ds/viewForums.do"
    }
  end
  
  # log in to the CollabNet instance
  def login
    @agent.get(@urls[:login]) do |page|
      page.forms[1].loginID  = @options[:username]
      page.forms[1].password = @options[:password]
      page.forms[1].click_button
    end
  end
  
  # return a hash of forums (text => url)
  def forums
    # go through all links on the forum index page looking for forum links
    @agent.get(@urls[:forums]).filtered_links('viewForumSummary').inject({}) do |hash, link|
      hash[link.text.strip] = link.href; hash
    end.reject { |key, value| @options[:forum] && key != @options[:forum] }
  end
  
  # return an array of messages for the given forum URL
  def messages(url)
    page     = @agent.get(url + '&viewType=thread&orderBy=createDate&orderType=asc')
    messages = page.filtered_links("viewMessage").collect { |link| link.href }
    
    while next_link = page.link_with(:text => "Next \302\273") do
      page = @agent.click next_link
      messages += page.filtered_links("viewMessage").collect { |link| link.href }
    end
    
    messages
  end
  
  # convert the message at url (CollabNet) to a mail message string
  def message_to_mail(url)
    message = {}
    
    @agent.get(url).search('table.axial tr').each do |row|
      unless row.search('th').first.nil?
        case row.search('th').first.content
        when 'Header'
          message[:headers] = row.search('td').first.to_html.extract_headers
          escaped_prefix = Regexp.escape(@options[:strip])
          message[:headers][:subject].gsub! /^((\s*(re|fwd?):\s*)*)#{escaped_prefix}\s*/i, '\1'
          message[:headers][:subject].strip!
        when 'Message'
          message[:body] = row.search('td').first.to_html.gsub /^<td>\s*|\s*<\/td>$/, ''
        end
      end
    end
    
    message
  end
end

# run the application
App.new.run
