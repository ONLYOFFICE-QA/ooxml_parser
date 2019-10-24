# frozen_string_literal: true

module OoxmlParser
  # Class for single relationship
  class Relationship < OOXMLDocumentObject
    # @return [String] id of relation
    attr_accessor :id
    # @return [String] type of relation
    attr_accessor :type
    # @return [String] target of relation
    attr_accessor :target

    # Parse Relationship
    # @param [Nokogiri::XML:Node] node with Relationship
    # @return [Relationship] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'Id'
          @id = value.value
        when 'Type'
          @type = value.value
        when 'Target'
          @target = value.value
        end
      end
      self
    end
  end
end
