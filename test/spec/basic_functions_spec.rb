$: << File.dirname(__FILE__) + '/../../lib/'
require 'kconv'
require 'viewpoint'
require 'json'

# To run this test put a file called 'creds.json' in this directory with the following format:
#   {"user":"myuser","pass":"mypass","endpoint":"https://mydomain.com/ews/exchange.asmx"}


describe "Test the basic features of Viewpoint" do
  before(:all) do
    creds = JSON.load(File.open("#{File.dirname(__FILE__)}/creds.json",'r'))
    @ews = Viewpoint::EWS::SOAP::ExchangeWebService.new
    @ews.endpoint = creds['endpoint']
    @ews.set_auth(creds['user'],creds['pass'])
  end

  it 'should retrieve the various Folder Types' do
    @ews.get_folder(:inbox).first.should
      be_instance_of Viewpoint::EWS::Folder
    @ews.get_folder(:calendar).first.should
      be_instance_of(Viewpoint::EWS::CalendarFolder)
    @ews.get_folder(:contacts).first.should
      be_instance_of(Viewpoint::EWS::ContactsFolder)
    @ews.get_folder(:tasks).first.should
      be_instance_of(Viewpoint::EWS::TasksFolder)
  end

  it 'should retrive the Inbox by name' do
    @ews.get_folder_by_name('Inbox').should
      be_instance_of(Viewpoint::EWS::Folder)
  end

  it 'should retrive the Inbox by FolderId' do
    inbox = (@ews.get_folder_by_name 'Inbox')
    @ews.get_folder(inbox.id).first.should
      be_instance_of(Viewpoint::EWS::Folder)
  end

  it 'should retrieve an Array of Folder Types' do
    folders = @ews.find_folder
    folders.should be_instance_of(Array)
    folders.first.should be_kind_of(Viewpoint::EWS::GenericFolder)
  end

  it 'should retrieve messages from a mail folder' do
    inbox = @ews.get_folder(:inbox).first
    msgs = inbox.find_items
    msgs.should be_instance_of(Array)
    if msgs.length > 0
      msgs.first.should be_kind_of(Viewpoint::EWS::Item)
    end
  end

  it 'should retrieve an item by id if one exists' do
    inbox = @ews.get_folder(:inbox).first
    msgs = inbox.find_items
    if msgs.length > 0
      item = inbox.get_item(msgs.first.id)
      item.should be_kind_of(Viewpoint::EWS::Item)
    else
      msgs.should be_empty
    end
  end

  it 'should retrieve a folder by name' do
    inbox = @ews.get_folder_by_name("Inbox")
    inbox.should be_instance_of(Viewpoint::EWS::Folder)
  end

  it 'should retrieve a list of folder names' do
    @ews.all_folders.should_not be_empty
  end

end
