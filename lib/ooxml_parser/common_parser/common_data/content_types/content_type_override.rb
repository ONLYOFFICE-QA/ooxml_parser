# frozen_string_literal: true

module OoxmlParser
  # Class for data from `Override` tag
  class ContentTypeOverride < OOXMLDocumentObject
    # @return [String] content type
    attr_reader :content_type
    # @return [String] part name
    attr_reader :part_name

    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ContentType'
          @content_type = value.value.to_s
        when 'PartName'
          @part_name = value.value.to_s
        end
      end
      self
    end
  end
end
