module OoxmlParser
  # Document Properties
  class DocumentProperties < OOXMLDocumentObject
    attr_accessor :pages, :words

    # Parse Document properties
    # @return [DocumentProperties]
    def parse
      properties_file = "#{OOXMLDocumentObject.path_to_folder}docProps/app.xml"
      unless File.exist?(properties_file)
        warn "There is no 'docProps/app.xml' in docx. It may be some problem with it"
        return self
      end
      node = Nokogiri::XML(File.open(properties_file))
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'Properties'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'Pages'
              @pages = node_child_child.text.to_i
            when 'Words'
              @words = node_child_child.text.to_i
            end
          end
        end
      end
      self
    end
  end
end
