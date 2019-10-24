# frozen_string_literal: true

module OoxmlParser
  # Class for data from `Default` tag
  class ContentTypeDefault < OOXMLDocumentObject
    # @return [String] content type
    attr_reader :content_type
    # @return [String] extension
    attr_reader :extension

    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ContentType'
          @content_type = value.value.to_s
        when 'Extension'
          @extension = value.value.to_s
        end
      end
      self
    end
  end
end
