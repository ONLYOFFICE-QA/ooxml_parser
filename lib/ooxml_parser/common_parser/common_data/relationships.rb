require_relative 'relationships/relationship'
module OoxmlParser
  # Class for describing list of relationships
  class Relationships
    # @return [Array, Relationship] array of relationships
    attr_accessor :relationship

    def initialize
      @relationship = []
    end

    # @return [Array, Column] accessor for relationship
    def [](key)
      @relationship[key]
    end

    # Parse Relationships
    # @param [Nokogiri::XML:Node] node with Relationships
    # @return [Relationships] result of parsing
    def self.parse(node)
      rels = Relationships.new
      node.xpath('*').each do |node_children|
        case node_children.name
        when 'Relationship'
          rels.relationship << Relationship.parse(node_children)
        end
      end
      rels
    end
  end
end
