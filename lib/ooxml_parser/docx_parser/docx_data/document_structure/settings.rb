module OoxmlParser
  # Class for parsing `settings.xml` file
  class Settings < OOXMLDocumentObject
    # @return [OoxmlSize] size of default tab
    attr_accessor :default_tab_stop

    # Parse Settings object
    # @return [Settings] result of parsing
    def parse
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/settings.xml'))
      doc.xpath('w:settings/*').each do |node_child|
        case node_child.name
        when 'defaultTabStop'
          @default_tab_stop = OoxmlSize.new.parse(node_child)
        end
      end
      self
    end
  end
end
