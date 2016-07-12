module OoxmlParser
  class LineJoin
    attr_accessor :type, :limit

    def self.parse(parent_node)
      line_join = LineJoin.new
      parent_node.xpath('*').each do |line_join_node|
        case line_join_node.name
        when 'round', 'bevel'
          line_join.type = line_join_node.name.to_sym
        when 'miter'
          line_join.type = :miter
          line_join.limit = line_join_node.attribute('lim').value.to_f
        end
      end
      line_join
    end
  end
end
