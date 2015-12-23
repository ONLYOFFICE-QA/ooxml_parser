module OoxmlParser
  class TargetElement
    attr_accessor :type, :id, :name, :built_in

    def initialize(type = '', id = '')
      @type = type
      @id = id
    end

    def self.parse(target_node)
      target = TargetElement.new
      target_node.xpath('*').each do |target_node_child|
        case target_node_child.name
        when 'sldTgt'
          target.type = :slide
        when 'sndTgt'
          target.type = :sound
          target.name = target_node_child.attribute('name').value
          target.built_in = target_node_child.attribute('builtIn').value ? StringHelper.to_bool(target_node_child.attribute('builtIn').value) : false
        when 'spTgt'
          target.type = :shape
          target.id = target_node_child.attribute('spid').value
        when 'inkTgt'
          target.type = :ink
          target.id = target_node_child.attribute('spid').value
        end
      end
      target
    end
  end
end
