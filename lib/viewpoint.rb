=begin
  This file is part of Viewpoint; the Ruby library for Microsoft Exchange Web Services.

  Copyright © 2011 Dan Wanek <dan.wanek@gmail.com>

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
=end

# We only what one instance of this class so include Singleton
require 'date'
require 'base64'
# bug in rubyntlm with ruby 1.9.x
require 'kconv' if(RUBY_VERSION.start_with? '1.9')


# Class Extensions
require 'extensions/string'

# Load the backend SOAP infrastructure.  Today this is Handsoap.
require 'soap/soap_provider'

# Load the model classes
# Base Models
require 'model/model'
require 'model/mailbox_user'
require 'model/attendee'
require 'model/generic_folder'
require 'model/item_field_uri_map' # supports Item
require 'model/item'

# Specific Models
# Folders
require 'model/folder'
require 'model/calendar_folder'
require 'model/contacts_folder'
require 'model/search_folder'
require 'model/tasks_folder'
# Items
require 'model/message'
require 'model/calendar_item'
require 'model/contact'
require 'model/distribution_list'
require 'model/meeting_message'
require 'model/meeting_request'
require 'model/meeting_response'
require 'model/meeting_cancellation'
require 'model/task'
require 'model/attachment'
require 'model/file_attachment'
require 'model/item_attachment'

require 'mail' # used to convert Message items to RFC822 compliant messages
require 'icalendar' # used to convert Calendar items to iCalendar objects

require 'exceptions/exceptions'

require 'helpers/autodiscover.rb'

# This is the class that controls access and presentation to
# Exchange Web Services.  It is possible to just use the SOAP classes
# themselves but this is what ties all the pieces together.
#
# @attr_reader [SOAP::ExchangeWebService] :ews The SOAP object used to make
#   calls to the Exchange Web Service.
module Viewpoint
  class InvalidCredentials < StandardError
    def initialize()
      @message = "Invalid credentials"
    end
  end

  class ExpiredCredentials < InvalidCredentials
    def initialize()
      @message = "Credentials have expired"
    end
  end

  class EndpointNotFound < StandardError
    def initialize()
      @message = "Unable to find EWS endpoint"
    end
  end
 
  class UnknownHttpError < StandardError
    def initialize(status_code)
      @message = "Unknown HTTP error (#{status_code})"
    end
  end

  module EWS
    class EwsBase
      def initialize(args = {})
        @ews = SOAP::ExchangeWebService.new
        @ews.set_auth(args[:user], args[:password])
        @ews.version = args[:version]
        @ews.endpoint = args[:endpoint] unless args[:endpoint].nil?
      end
    end

    class EmailAccount < EwsBase
      def mailboxes
        @ews.all_folders
      end
    end

    # @attr_reader [Viewpoint::EWS::SOAP::ExchangeWebService] :ews The EWS
    #   object used to make SOAP calls. You typically don't need to use this,
    #   but if you want to play around with the SOAP back-end it's available.
    class EWS
      include Viewpoint

      attr_reader :endpoint
      attr_reader :ews, :user, :password

      def initialize(opts = {})
        @ews = SOAP::ExchangeWebService.new(opts)
      end

      # Set the endpoint for Exchange Web Services.  
      # @param [String] endpoint The URL of the endpoint. This should end in
      #   'exchange.asmx' and is typically something like this:
      #   https://myexch/ews/exchange.asmx
      # @param [Integer] version The SOAP version to use.  This defaults to 1
      #   and you should not need to pass this parameter.
      def endpoint=(endpoint, version = 1)
        @ews.endpoint = endpoint
      end

      # Set the SOAP username and password.
      # @param [String] user The user name
      # @param [String] pass The password
      def set_auth(user,pass)
        @ews.set_auth(user, pass)
      end

      # Set the http driver that the SOAP back-end will use.
      # @param [Symbol] driver The HTTP driver.  Available drivers:
      #   :curb, :net_http, :http_client(Default)
      def set_http_driver(driver)
        Handsoap.http_driver = driver
      end

      # Sets the CA path to a certificate or hashed certificate directory.
      # This is the same as HTTPClient::SSLConfig.set_trust_ca
      # @param [String] ca_path A path to an OpenSSL::X509::Certificate or a
      #     'c-rehash'ed directory
      def set_trust_ca(ca_path)
        SOAP::ExchangeWebService.set_http_options(:trust_ca_file => ca_path)
      end

      #
      # Utility method to return all folders that contain email. By default
      # it won't return the following:
      #     "Junk Email"
      #     "Deleted Items"
      #     "Conversation Action Settings"
      # To include those folders in the returned array specify any of the
      # following options:
      #     :return_junk_email => true
      #     :return_deleted_items => true
      #     :return_conversation_action_settings => true
      # To specifically exclude a list of folders specify:
      #     :ignore_folders [array of folders to exclude]
      def get_all_email_folders(opts = {})
        ignore_folders = opts[:ignore_folders] || []
        ignore_folders.push "conversation action setttings" unless
          opts[:conversation_action_settings]
        ignore_folders.push "deleted items" unless
          opts[:deleted_items]
        ignore_folders.push "junk email" unless
          opts[:junk_email]

        email_folders = []
        @ews.all_folders.each do |f|
          next unless f.class == Viewpoint::EWS::Folder
          next if ignore_folders.include? f.display_name.downcase

          email_folders.push f
        end

        email_folders
      end
      
      #
      # Utility method to return all folders that contain contacts. By default
      # To specifically exclude a list of folders specify:
      #     :ignore_folders [array of folders to exclude]
      def get_all_contacts_folders(opts = {})
        ignore_folders = opts[:ignore_folders] || []

        contacts_folders = []
        @ews.all_folders.each do |f|
          next unless f.class == Viewpoint::EWS::ContactsFolder
          next if ignore_folders.include? f.display_name.downcase

          contacts_folders.push f
        end

        contacts_folders
      end

      def get_folder(folder_ids)
        @ews.get_folder(folder_ids)
      end

      def get_contacts
        @ews.get_folder(:contacts).each{ |f| f.find_items }
      end

      def get_item(item_id, item_shape = nil)
        @ews.get_item([item_id].flatten, item_shape)
      end
    end # class EWS
  end # module EWS
end
