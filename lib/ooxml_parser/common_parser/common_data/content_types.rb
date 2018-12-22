require_relative 'content_types/content_type_default'
require_relative 'content_types/content_type_override'
module OoxmlParser
  # Class for data from `[Content_Types].xml` file
  class ContentTypes < OOXMLDocumentObject
    # @return [Array] list of content types
    attr_accessor :content_types_list

    def initialize(parent: nil)
      @parent = parent
      @content_types_list = []
    end

    # @return [Array] accessor
    def [](key)
      @content_types_list[key]
    end

    def parse
      doc = Nokogiri::XML.parse(File.open("#{OOXMLDocumentObject.path_to_folder}/[Content_Types].xml"))
      node = doc.xpath('*').first

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'Default'
          @content_types_list << ContentTypeDefault.new(parent: self).parse(node_child)
        when 'Override'
          @content_types_list << ContentTypeOverride.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
