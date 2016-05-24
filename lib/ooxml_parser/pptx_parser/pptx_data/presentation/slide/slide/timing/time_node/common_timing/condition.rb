module OoxmlParser
  class Condition
    attr_accessor :event, :delay, :duration

    def initialize(event = nil, delay = nil, duration = nil)
      @event = event
      @delay = delay
      @duration = duration
    end

    def self.parse(condition_node)
      condition = Condition.new
      condition.event = condition_node.attribute('evt').value if condition_node.attribute('evt')
      condition.delay = condition_node.attribute('delay').value if condition_node.attribute('delay')
      condition
    end

    def self.parse_list(conditions_list_node)
      conditions = []
      conditions_list_node.xpath('p:cond').each do |condition_node|
        conditions << Condition.parse(condition_node)
      end
      conditions
    end
  end
end
