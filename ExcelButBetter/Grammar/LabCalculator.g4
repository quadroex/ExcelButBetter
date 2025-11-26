grammar LabCalculator;

options {
    language=CSharp;
}

@lexer::namespace { ExcelButBetter.Grammar }
@parser::namespace { ExcelButBetter.Grammar }

/* --- ПРАВИЛА ПАРСЕРА (Синтаксис) --- */

// Головна точка входу: вираз і кінець файлу (EOF)
compileUnit
    : expression EOF
    ;

expression
    // 1. Функції (найвищий пріоритет)
    : INC LPAREN expression RPAREN       # IncExpr

| DEC LPAREN expression RPAREN       # DecExpr
    
    // 2. Унарні операції (знак числа та логічне НЕ)

| NOT expression                     # NotExpr
| PLUS expression                    # UnaryPlusExpr
| MINUS expression                   # UnaryMinusExpr
    
    // 3. Множення та ділення

| expression (MULTIPLY | DIVIDE) expression  # MultiplicativeExpr
    
    // 4. Додавання та віднімання

| expression (PLUS | MINUS) expression       # AdditiveExpr
    
    // 5. Оператори порівняння (найнижчий пріоритет)

| expression (EQ | LT | GT) expression       # RelationalExpr
    
    // 6. Базові елементи (атоми)

| NUMBER                             # NumberExpr
| IDENTIFIER                         # IdentifierExpr
| LPAREN expression RPAREN           # ParenthesizedExpr
    ;

/* --- ПРАВИЛА ЛЕКСЕРА (Слова) --- */

// Ключові слова (case-insensitive - нечутливі до регістру, але в g4 простіше писати так)
INC: 'inc';
DEC: 'dec';
NOT: 'not';

// Арифметика
PLUS: '+';
MINUS: '-';
MULTIPLY: '*';
DIVIDE: '/';

// Порівняння
EQ: '=';
LT: '<';
GT: '>';

// Розділові знаки
LPAREN: '(';
RPAREN: ')';

// Числа: цілі (5) або дробові (5.23)
NUMBER: [0-9]+ ('.' [0-9]+)?;

// Ідентифікатори клітинок: Літери + Цифри (наприклад, A1, B52, AB10)
IDENTIFIER: [A-Z]+ [0-9]+;

// Пропускаємо пробіли та переноси рядків
WS: [ \t\r\n]+ -> skip;