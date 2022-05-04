# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w15:commentEx` tag
  class CommentExtended < OOXMLDocumentObject
    # @return [Integer] id of paragraph
    attr_accessor :paragraph_id
    # @return [True, False] is done?
    attr_accessor :done

    # Parse CommentExtended object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommentExtended] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'paraId'
          @paragraph_id = value.value.to_i
        when 'done'
          @done = attribute_enabled?(value.value)
        end
      end
      self
    end
  end
end
