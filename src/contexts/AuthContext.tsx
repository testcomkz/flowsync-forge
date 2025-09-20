import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { SupabaseService } from '@/services/supabaseService';

const supabaseService = new SupabaseService();

interface User {
  id: number;
  email: string;
  full_name: string;
  role: string;
}

interface AuthContextType {
  user: User | null;
  login: (email: string, password: string) => Promise<boolean>;
  logout: () => void;
  isAuthenticated: boolean;
  isLoading: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
};

interface AuthProviderProps {
  children: ReactNode;
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  // Восстанавливаем состояние пользователя при загрузке
  useEffect(() => {
    const restoreUserSession = () => {
      try {
        const savedUser = localStorage.getItem('auth_user');
        if (savedUser) {
          const userData = JSON.parse(savedUser);
          setUser(userData);
          console.log('✅ User session restored from localStorage');
        }
      } catch (error) {
        console.error('Error restoring user session:', error);
        localStorage.removeItem('auth_user');
      } finally {
        setIsLoading(false);
      }
    };

    restoreUserSession();
  }, []);

  const login = async (email: string, password: string): Promise<boolean> => {
    try {
      const result = await supabaseService.authenticateUser(email, password);
      
      if (result.success && result.user) {
        const userData = {
          id: result.user.id,
          email: result.user.email,
          full_name: result.user.full_name,
          role: result.user.role
        };
        setUser(userData);
        
        // Сохраняем в localStorage
        localStorage.setItem('auth_user', JSON.stringify(userData));
        console.log('✅ User session saved to localStorage');
        return true;
      } else {
        alert(result.error || 'Ошибка входа');
        return false;
      }
    } catch (error) {
      alert('Ошибка подключения к серверу');
      return false;
    }
  };

  const logout = () => {
    setUser(null);
    localStorage.removeItem('auth_user');
    console.log('✅ User session cleared from localStorage');
  };

  const value = {
    user,
    login,
    logout,
    isAuthenticated: !!user,
    isLoading
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};
