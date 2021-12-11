from pprint import pprint

from PyInquirer import prompt, Separator

from examples import custom_style_1


questions = [
    {
        'type': 'checkbox',
        'qmark': '',
        'message': 'Select toppings',
        'name': 'toppings',
        'choices': [
            {
                'name': 'Pineapple'
            },
            {
                'name': 'Olives',
            },
            {
                'name': 'Extra cheese'
            }
        ],
        'validate': lambda answer: 'You must choose at least one topping.'
        if len(answer) == 0 else True
    }
]

answers = prompt(questions, style=custom_style_1)
pprint(answers)
