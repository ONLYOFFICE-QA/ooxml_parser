require_relative 'relationships/relationship'
module OoxmlParser
  # Class for describing list of relationships
  class Relationships
    # @return [Array, Relationship] array of relationships
    attr_accessor :relationship

    def initialize(parent: nil)
      @relationship = []
      @parent = parent
    end

    # @return [Array, Column] accessor for relationship
    def [](key)
      @relationship[key]
    end

    # Parse Relationships
    # @param [Nokogiri::XML:Node] node with Relationships
    # @return [Relationships] result of parsing
    def parse(node)
      node.xpath('*').each do |node_children|
        case node_children.name
        when 'Relationship'
          @relationship << Relationship.new(parent: self).parse(node_children)
        end
      end
      self
    end

    # Parse .rels file
    # @param file_path [String] path to file
    # @return [Relationships]
    def parse_file(file_path)
      node = parse_xml(file_path)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'Relationships'
          parse(node_child)
        end
      end
      self
    end

    # Get target name by id
    # @param id [String] id of target
    # @return [String] target name
    def target_by_id(id)
      @relationship.each do |cur_rels|
        return cur_rels.target if cur_rels.id == id
      end
      nil
    end

    # Get target name by type
    # @param type [String] type of target
    # @return [String] target name
    def target_by_type(type)
      @relationship.each do |cur_rels|
        return cur_rels.target if cur_rels.type.include?(type)
      end
      nil
    end
  end
end
