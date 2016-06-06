module OoxmlParser
  # Class for single relationship
  class Relationship
    # @return [String] id of relation
    attr_accessor :id
    # @return [String] type of relation
    attr_accessor :type
    # @return [String] target of relation
    attr_accessor :target

    # Parse Relationship
    # @param [Nokogiri::XML:Node] node with Relationship
    # @return [Relationship] result of parsing
    def self.parse(node)
      rel = Relationship.new
      rel.id = node.attribute('Id').value
      rel.type = node.attribute('Type').value
      rel.target = node.attribute('Target').value
      rel
    end
  end
end
