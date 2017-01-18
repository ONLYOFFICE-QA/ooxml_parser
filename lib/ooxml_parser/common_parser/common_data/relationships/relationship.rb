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
      @id = node.attribute('Id').value
      @type = node.attribute('Type').value
      @target = node.attribute('Target').value
      self
    end
  end
end
