grammar LabCalculator;

options {
    language=CSharp;
}

@lexer::namespace { ExcelButBetter.Grammar }
@parser::namespace { ExcelButBetter.Grammar }

compileUnit
    : expression EOF
    ;

expression
    : INC LPAREN expression RPAREN       # IncExpr

| DEC LPAREN expression RPAREN       # DecExpr

| NOT expression                     # NotExpr
| PLUS expression                    # UnaryPlusExpr
| MINUS expression                   # UnaryMinusExpr

| expression (MULTIPLY | DIVIDE) expression  # MultiplicativeExpr

| expression (PLUS | MINUS) expression       # AdditiveExpr

| expression (EQ | LT | GT) expression       # RelationalExpr

| NUMBER                             # NumberExpr
| IDENTIFIER                         # IdentifierExpr
| LPAREN expression RPAREN           # ParenthesizedExpr
    ;

INC: 'inc';
DEC: 'dec';
NOT: 'not';

PLUS: '+';
MINUS: '-';
MULTIPLY: '*';
DIVIDE: '/';

EQ: '=';
LT: '<';
GT: '>';

LPAREN: '(';
RPAREN: ')';

NUMBER: [0-9]+ ('.' [0-9]+)?;

IDENTIFIER: [A-Z]+ [0-9]+;

WS: [ \t\r\n]+ -> skip;
