require_relative 'style_sheet/number_formats'
module OoxmlParser
  # Parsing file styles.xml
  class StyleSheet < OOXMLDocumentObject
    attr_accessor :number_formats
    def parse
      doc = Nokogiri::XML(File.open("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/styles.xml"))
      doc.root.xpath('*').each do |node_child|
        case node_child.name
        when 'numFmts'
          @number_formats = NumberFormats.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
