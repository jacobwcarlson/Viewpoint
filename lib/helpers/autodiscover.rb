require 'nokogiri'
require 'resolv'
require 'httpclient'

module Viewpoint
  module EWS
    module Autodiscover
      SCHEMA_PREFIX= "http://schemas.microsoft.com/exchange/autodiscover"
      REQUEST_SCHEMA_URL = "#{SCHEMA_PREFIX}/outlook/requestschema/2006"
      RESPONSE_SCHEMA_URL = "#{SCHEMA_PREFIX}/outlook/responseschema/2006a"
      MAX_REDIRECTS = 10

      NAMESPACES = {
        'a' => "#{SCHEMA_PREFIX}/responseschema/2006",
        'o' => "#{SCHEMA_PREFIX}/outlook/responseschema/2006a"
      }

      E_EXPR_PROTOCOL = 'o:Protocol[o:Type="EXPR"]'
      E_EWS_URL = 'o:EwsUrl'
      E_URL_TTL = 'o:TTL'
      E_REDIRECT_URL = 'o:RedirectUrl'
      E_ACTION = 'o:Action'
      E_ACCOUNT = 'a:Autodiscover/o:Response/o:Account'

      class Redirect < StandardError
        attr_reader :url
        def initialize(args = {})
          @url = args[:url]
        end
      end

      class ReAddr < StandardError
        attr_reader :email_address
        def initialize(args = {})
          @email_address = args[:email_address]
        end
      end

      # Find the EWS endpoint via Autodiscover
      def self.find_endpoint(email_address, password)
        raise ArgumentError, "Missing email address" unless email_address

        @redirect_count = 0

        # Build the XML request
        request_body = Nokogiri::XML::Builder.new do |xml| 
          xml.Autodiscover('xmlns' => REQUEST_SCHEMA_URL){
            xml.Request {
              xml.EMailAddress email_address
              xml.AcceptableResponseSchema RESPONSE_SCHEMA_URL
            }
          }
        end.to_xml

        # Create a list of likely URLs for the autodiscover service.
        urls = []
        ad_path = "/autodiscover/autodiscover.xml"
        domain = email_address.sub(/^.+@/, "")
        urls.push "https://#{domain}#{ad_path}"
        urls.push "https://autodiscover.#{domain}#{ad_path}"
        urls.push "http://autodiscover.#{domain}#{ad_path}"

        response = nil
        opts = { 'Content-Type' => 'text/xml; charset=utf-8'}
        redirect_count = 0
        http = HTTPClient.new
        urls.each do |url|
          while redirect_count < MAX_REDIRECTS do 
            http.set_auth(url, email_address, password)
            begin
              response = http.post(url, request_body, opts)
            rescue
              break
            end

            begin
              return parse_autodiscover_response response
            rescue Redirect => e
              redirect_count += 1
              url = e.url
              next
            # We don't actually do anything here as I'm not really sure how to
            # proceed. Just log it an return failure.
            rescue ReAddr => e
              $stderr.printf "Asked to switch email addresses from %s to %s\n",
                email_address, e.email_address
              $stderr.printf "But I don't know how to do that :(.\n"
              return nil
            rescue
              puts "Unknown exception: #{$!}"
            end
            
            break
          end
        end
      end

    private
      def self.parse_autodiscover_response(response)
        if response.status_code == 302
          raise Redirect.new(:url => response.header['Location'].first)
        end

        doc = Nokogiri::XML(response.content)
        account = doc.at_xpath(E_ACCOUNT, NAMESPACES)
        return nil unless account
        action = account.at_xpath(E_ACTION, NAMESPACES)
        return nil unless action

        case action.content
        # Server wants us to talk to another Autodiscover service.
        when 'redirectUrl'
          url = account.at_xpath(E_REDIRECT_URL, NAMESPACES)
          raise Redirect(:url => url.content)
        # We're supposed to replace the supplied email address with something
        when 'reAddr'
          raise ReAddr.new(:email_address => "bob@bob.com")
        # This is what we were hoping to find.
        when 'settings'
          protocol = account.at_xpath(E_EXPR_PROTOCOL, NAMESPACES)
          return nil unless protocol

          ews_url = protocol.at_xpath(E_EWS_URL, NAMESPACES)
          return nil if not ews_url
          # The default TTL, if not specifiedin the response, is 1 hour.
          # XXX: We should do something with the TTL and we don't
          #ttl = protocol.at_xpath(E_URL_TTL, NAMESPACES) || 1

          return ews_url.content
        end
      end
    end # module Autodiscover
  end # module EWS
end # module Viewpoint

